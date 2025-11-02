/*****************************************************
 * NOVAMAIL — BLOC 1 : CORE
 * Versions, constantes, helpers, storage utilities
 *****************************************************/

/* ====== CONFIGURATION / VERSION ====== */
// Retourne la version active (pour tests tu peux modifier ici
// mais en production ceci sera injecté/paramétré automatiquement).
function getAppVersion() {
  return "PRO"; // FREE | STARTER | PRO | BUSINESS
}

function getVersionConfig() {
  var version = (getAppVersion() || "FREE").toString().toUpperCase();

  var configs = {
    FREE: {
      name: "Free",
      maxRecipients: 10, // Par campagne
      monthlyQuota: 50, // Max destinataires par mois
      annualQuota: 500, // Max destinataires par an
      allowAttachments: false,
      allowImportSheets: false,
      allowTemplateSave: false,
      multiSender: false,
      scheduleSend: false,
      customBranding: false,
    },
    STARTER: {
      name: "Starter",
      maxRecipients: 200,
      monthlyQuota: 2000,
      annualQuota: 20000,
      allowAttachments: true,
      allowImportSheets: true,
      allowTemplateSave: true,
      multiSender: false,
      scheduleSend: "limited",
      customBranding: false,
    },
    PRO: {
      name: "Pro",
      maxRecipients: 1000,
      monthlyQuota: 10000,
      annualQuota: 120000,
      allowAttachments: true,
      allowImportSheets: true,
      allowTemplateSave: true,
      multiSender: true,
      scheduleSend: true,
      customBranding: true,
    },
    BUSINESS: {
      name: "Business",
      maxRecipients: 5000,
      monthlyQuota: 50000,
      annualQuota: 600000,
      allowAttachments: true,
      allowImportSheets: true,
      allowTemplateSave: true,
      multiSender: true,
      scheduleSend: "recurring",
      customBranding: true,
    },
  };

  return configs[version] || configs["FREE"];
}
//Fonctions associées pour gérer les quotas
function getUserQuota() {
  var raw = PropertiesService.getUserProperties().getProperty("USER_QUOTA");
  if (!raw)
    return { monthly: 0, annual: 0, lastReset: new Date().toISOString() };
  return JSON.parse(raw);
}

//Sauvegarder les quotas
function saveUserQuota(quota) {
  PropertiesService.getUserProperties().setProperty(
    "USER_QUOTA",
    JSON.stringify(quota)
  );
}

/* ====== DEFAULTS ====== */
var PRO_EMAIL = "foreverjoyfulcreations@gmail.com";
var SENDER_NAME = "NovaMail";
var BATCH_SIZE = 40;

/*****************************************************
 * NOVAMAIL — GLOBAL CONSTANTES
 *****************************************************/
var SCHEDULED_PREFIX = "SCHEDULED_"; // var = hoistée (évite le bug "before initialization")
const HISTORY_SHEET_NAME = "Historique_Campagnes";
const HISTORY_FILE_NAME = "NovaMail_Historique"; // nom unique du fichier
const HISTORY_PROP_KEY = "HISTORY_SPREADSHEET_ID";
const SCRIPT_PROPS = PropertiesService.getScriptProperties();

/**
 * Crée un objet campagne standardisé
 */
function buildCampaignData(
  type,
  recipients,
  subject,
  htmlBody,
  attachments,
  sendAt
) {
  return {
    type: type, // "directe" ou "planifiée"
    name: type === "directe" ? "Campagne Directe" : "Campagne Planifiée",
    subject: subject || "",
    htmlBody: htmlBody || "",
    attachments: attachments || [],
    recipients: recipients || [],
    sendAt: sendAt || null,
    createdAt: new Date().toISOString(),
  };
}

// ============================================================
// 🧭 OUTILS GÉNÉRAUX
// ============================================================
/**
 * 🗓️ Format de date lisible en français
 */
function formatDateFR(date) {
  const opts = {
    day: "2-digit",
    month: "long",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  };
  return date.toLocaleString("fr-FR", opts);
}

/** Journalisation simplifiée */
function log(msg) {
  Logger.log(`[NovaMail] ${msg}`);
}

/** Nettoie les erreurs renvoyées vers le front */
function handleError(err) {
  log("❌ " + err.message);
  throw new Error(err.message);
}

/* ====== KEYS ====== */
var USER_PROP_PREFIX = "NOVAMAIL_USER"; // per-user storage
var SCRIPT_PROP_PREFIX = "NOVAMAIL_SCRIPT"; // global per-script storage
var USER_KEYS = {
  LAST_LIST: "LAST_LIST",
  TEMPLATE_PREFIX: "TEMPLATE:", // full key = USER_PROP_PREFIX + ":" + TEMPLATE_PREFIX + name
};

/* ====== UTILITAIRES ====== */

function normalizeEmail(email) {
  if (!email) return "";
  return String(email).trim().toLowerCase();
}

function isValidEmail(email) {
  var re = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
  return re.test(normalizeEmail(email));
}

/**
 * parseRecipientsFromText(text)
 * Accepte lines "Nom,Email" ou "Email" par ligne.
 * Retourne [{nom,email}, ...]
 */
function parseRecipientsFromText(text) {
  if (!text) return [];
  var lines = String(text)
    .split(/\r?\n/)
    .map(function (l) {
      return l.trim();
    })
    .filter(Boolean);
  var out = [];
  lines.forEach(function (line) {
    var parts = line.split(",");
    var name = "";
    var email = "";
    if (parts.length >= 2) {
      name = parts[0].trim();
      email = parts[1].trim();
    } else {
      email = parts[0].trim();
    }
    email = normalizeEmail(email);
    if (!isValidEmail(email)) return;
    if (!name) {
      name = email.split("@")[0].replace(/[._-]+/g, " ");
      name = name.replace(/\b\w/g, function (c) {
        return c.toUpperCase();
      });
    }
    out.push({ nom: name, email: email });
  });
  return out;
}

/* ====== USER PROPERTIES HELPERS ====== */

function saveUserProp(key, obj) {
  if (!key) return;
  var store = PropertiesService.getUserProperties();
  store.setProperty(USER_PROP_PREFIX + ":" + key, JSON.stringify(obj));
}

function loadUserProp(key) {
  var raw = PropertiesService.getUserProperties().getProperty(
    USER_PROP_PREFIX + ":" + key
  );
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

/* save last imported list (per-user) */
function saveLastImportedList(list) {
  if (!list || !list.length) return false;
  saveUserProp(USER_KEYS.LAST_LIST, list);
  return true;
}

/* ====== SCRIPT PROPERTIES HELPERS (global) used for scheduling jobs ====== */
function setScriptProp(key, valueObj) {
  PropertiesService.getScriptProperties().setProperty(
    SCRIPT_PROP_PREFIX + ":" + key,
    JSON.stringify(valueObj)
  );
}
function getScriptProp(key) {
  var raw = PropertiesService.getScriptProperties().getProperty(
    SCRIPT_PROP_PREFIX + ":" + key
  );
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}
function deleteScriptProp(key) {
  PropertiesService.getScriptProperties().deleteProperty(
    SCRIPT_PROP_PREFIX + ":" + key
  );
}
function listAllScriptProps() {
  var props = PropertiesService.getScriptProperties().getProperties();
  var out = {};
  for (var k in props) {
    if (k.indexOf(SCRIPT_PROP_PREFIX + ":") === 0) {
      out[k.substring((SCRIPT_PROP_PREFIX + ":").length)] = props[k];
    }
  }
  return out;
}

/*****************************************************
 * NOVAMAIL — BLOC 2 : IMPORTS & TEMPLATES
 *****************************************************/

/**
 * listGoogleSheets()
 * Retourne une liste simple des Google Sheets accessibles.
 * (Utilisé par le front pour afficher un select)
 */
function listGoogleSheets() {
  var files = DriveApp.searchFiles(
    "mimeType='application/vnd.google-apps.spreadsheet'"
  );
  var out = [];
  while (files.hasNext()) {
    var f = files.next();
    out.push({
      id: f.getId(),
      name: f.getName(),
      modifiedDate: f.getLastUpdated() ? f.getLastUpdated().toISOString() : "",
    });
  }
  return out;
}

/**
 * importFromSheet(fileId, sheetName)
 * Lit la feuille et retourne [{nom,email},...]
 */
function importFromSheet(fileId, sheetName) {
  if (!fileId) throw new Error("importFromSheet: fileId manquant.");
  var ss = SpreadsheetApp.openById(fileId);
  var sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
  if (!sheet) throw new Error("importFromSheet: feuille introuvable.");

  var data = sheet.getDataRange().getValues();
  if (!data || !data.length) return [];

  var headers = data[0].map(function (h) {
    return String(h).toLowerCase().trim();
  });
  var nameIndex = -1,
    emailIndex = -1;
  headers.forEach(function (h, i) {
    if (h.indexOf("nom") !== -1 || h.indexOf("name") !== -1) nameIndex = i;
    if (h.indexOf("mail") !== -1 || h.indexOf("email") !== -1) emailIndex = i;
  });
  if (emailIndex === -1)
    throw new Error("importFromSheet: colonne 'email' introuvable.");

  var list = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var email = normalizeEmail(row[emailIndex]);
    if (!isValidEmail(email)) continue;
    var name =
      nameIndex !== -1 && row[nameIndex]
        ? String(row[nameIndex])
        : email.split("@")[0];
    list.push({ nom: name, email: email });
  }

  // sauvegarde en tant que dernière liste
  try {
    saveLastImportedList(list);
  } catch (e) {
    /* ignore */
  }

  return list;
}

/**
 * 📂 Parse un fichier CSV uploadé depuis l’interface Web.
 * (Le front-end enverra le fichier en Base64 à cette fonction)
 */
function parseCsvFile(base64Data) {
  if (!base64Data) throw new Error("Aucun contenu reçu.");
  var csvText = Utilities.newBlob(
    Utilities.base64Decode(base64Data)
  ).getDataAsString();
  return convertCsvToRecipients(csvText);
}

/**
 * convertCsvToRecipients(csvText)
 * Convertit texte CSV en tableau {nom,email}
 * Le front envoie le texte CSV brut (pas base64).
 */
function convertCsvToRecipients(csvText) {
  if (!csvText) return [];
  // Support simple CSV : ligne par ligne, champs séparés par virgule
  var lines = csvText
    .split(/\r?\n/)
    .map(function (l) {
      return l.trim();
    })
    .filter(Boolean);
  var list = [];
  lines.forEach(function (line) {
    var parts = line.split(",");
    var name = "",
      email = "";
    if (parts.length >= 2) {
      name = parts[0].trim();
      email = parts[1].trim();
    } else {
      email = parts[0].trim();
    }
    email = normalizeEmail(email);
    if (!isValidEmail(email)) return;
    if (!name) name = email.split("@")[0];
    list.push({ nom: name, email: email });
  });
  // save last list
  try {
    saveLastImportedList(list);
  } catch (e) {
    /* ignore */
  }
  return list;
}

/** Wrapper exposée au front (nom attendu) */
function importCsv(csvContent) {
  // Le client envoie le texte CSV brut (reader.readAsText)
  return convertCsvToRecipients(csvContent);
}

/** ===============================
 * 1️⃣ saveLastImportedList
 * Sauvegarde la dernière liste de destinataires dans les propriétés utilisateur.
 * Utile pour que la fonction getLastImportedList() puisse récupérer la dernière importation.
 * @param {Array} list - Tableau d’objets {name, email}
 */
function saveLastImportedList(list) {
  if (!list || !list.length) return; // rien à sauvegarder
  PropertiesService.getUserProperties().setProperty(
    USER_KEYS.LAST_LIST,
    JSON.stringify(list)
  );
}

/** ===============================
 * 2️⃣ getLastImportedList
 * Récupère la dernière liste de destinataires sauvegardée.
 * @returns {Array} Tableau d’objets {name, email} ou tableau vide si aucune liste.
 */
function getLastImportedList() {
  var raw = PropertiesService.getUserProperties().getProperty(
    USER_KEYS.LAST_LIST
  );
  return raw ? JSON.parse(raw) : [];
}

/** ===============================
 * 3️⃣ clearLastImportedList
 * Supprime la dernière liste de destinataires sauvegardée.
 */
function clearLastImportedList() {
  PropertiesService.getUserProperties().deleteProperty(USER_KEYS.LAST_LIST);
}

/** Prévisualisation rapide */
function previewEmail(nom, email, subject, html) {
  return {
    subject,
    htmlBody: html
      .replace(/{{\s*nom\s*}}/gi, nom || "Ami(e)")
      .replace(/{{\s*email\s*}}/gi, email),
  };
}

/* ===== Templates: sauvegarde / lecture ===== */

/**
 * saveTemplate(name, subject, htmlBody)
 * Sauvegarde le template pour l'utilisateur courant.
 */
function saveTemplate(name, subject, htmlBody) {
  if (!name) throw new Error("saveTemplate: nom requis.");
  var obj = {
    name: name,
    subject: subject || "",
    htmlBody: htmlBody || "",
    updatedAt: new Date().toISOString(),
  };
  // clé per-user: USER_PROP_PREFIX + ":" + TEMPLATE_PREFIX + name
  PropertiesService.getUserProperties().setProperty(
    USER_PROP_PREFIX + ":" + USER_KEYS.TEMPLATE_PREFIX + name,
    JSON.stringify(obj)
  );
  return obj;
}

/**
 * loadTemplate(name)
 * Charge le template par nom (retourne null si absent)
 */
function loadTemplate(name) {
  if (!name) throw new Error("loadTemplate: nom requis.");
  var raw = PropertiesService.getUserProperties().getProperty(
    USER_PROP_PREFIX + ":" + USER_KEYS.TEMPLATE_PREFIX + name
  );
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

/* Fournir une méthode pour lister templates (utile pour UI ultérieure) */
function listTemplates() {
  var props = PropertiesService.getUserProperties().getProperties();
  var arr = [];
  for (var k in props) {
    if (k.indexOf(USER_PROP_PREFIX + ":" + USER_KEYS.TEMPLATE_PREFIX) === 0) {
      try {
        arr.push(JSON.parse(props[k]));
      } catch (e) {
        /* ignore */
      }
    }
  }
  // tri par date
  arr.sort(function (a, b) {
    return (b.updatedAt || "").localeCompare(a.updatedAt || "");
  });
  return arr;
}

/**
 * 🗑 Supprime un template précis par nom
 */
function deleteTemplate(name) {
  if (!name) throw new Error("deleteTemplate: nom requis.");

  const key = USER_PROP_PREFIX + ":" + USER_KEYS.TEMPLATE_PREFIX + name;
  const props = PropertiesService.getUserProperties();

  if (!props.getProperty(key)) {
    throw new Error("Modèle introuvable : " + name);
  }

  props.deleteProperty(key);
  Logger.log("🗑 Modèle supprimé : " + name);
  return { success: true, name };
}

/**
 * 🧹 Supprime tous les templates de l’utilisateur courant
 */
function deleteAllTemplates() {
  const props = PropertiesService.getUserProperties();
  const all = props.getProperties();
  let count = 0;

  for (let key in all) {
    if (key.startsWith(USER_PROP_PREFIX + ":" + USER_KEYS.TEMPLATE_PREFIX)) {
      props.deleteProperty(key);
      count++;
    }
  }

  Logger.log(`🧹 ${count} modèle(s) supprimé(s) pour l’utilisateur courant`);
  return { success: true, count };
}

/**
 * Upload d'une pièce jointe (base64 string) -> création d'un fichier Drive temporaire.
 * Retourne l'ID du fichier Drive.
 *
 * ATTENTION: Conserver peu de fichiers, supprimer après envoi si nécessaire.
 */
function _createTempFileFromBase64(base64Str, filename, mimeType) {
  // Décoder base64 -> blob
  var data = Utilities.base64Decode(base64Str);
  var blob = Utilities.newBlob(
    data,
    mimeType || "application/octet-stream",
    filename
  );
  // Créer dans Drive (dans ton Drive racine)
  var file = DriveApp.createFile(blob);
  // Optionnel : placer dans un dossier spécifique ou tagger via Properties
  return file.getId();
}

/** ===============================
 * 4️⃣ sendTestEmail
 * Envoie un email de test avec les mêmes options que sendCampaign.
 * Permet de vérifier le rendu HTML, les pièces jointes et le sujet.
 * @param {string} to - Adresse email destinataire du test
 * @param {string} subject - Sujet de l’email
 * @param {string} htmlBody - Contenu HTML de l’email
 * @param {Array} attachments - Tableau de fichiers {name, base64, mime}
 * @returns {string} Message de confirmation
 */
/*****************************************************
 * 📬 Envoi d’un email test au propriétaire du script
 * (robuste, logué, avec historique)
 *****************************************************/
/**
 * 🚀 Envoi d’un test à l’utilisateur courant
 * Stocké aussi dans l'historique
 */
function sendTestToMe(subject, htmlBody, attachments) {
  try {
    // 🔹 Détermination de l’adresse de destination
    const me =
      Session.getActiveUser().getEmail() ||
      Session.getEffectiveUser().getEmail();

    if (!me || me.indexOf("@") === -1)
      throw new Error("Aucun email valide détecté pour l’utilisateur actif.");

    // 🔹 Conversion des pièces jointes (si présentes)
    const blobs = [];
    if (attachments && attachments.length) {
      attachments.forEach((att) => {
        try {
          const data = Utilities.base64Decode(att.base64);
          const blob = Utilities.newBlob(
            data,
            att.mime || "application/octet-stream",
            att.name
          );
          blobs.push(blob);
        } catch (err) {
          Logger.log(
            `⚠️ Erreur décodage pièce jointe ${att.name}: ${err.message}`
          );
        }
      });
    }

    // 🔹 Construction des options d’envoi Gmail
    const mailOptions = {
      htmlBody: htmlBody,
      name: SENDER_NAME || "Campagne Test",
      replyTo: PRO_EMAIL || me,
    };
    if (blobs.length) mailOptions.attachments = blobs;

    // 🔹 Envoi réel
    GmailApp.sendEmail(
      me,
      subject || "(test)",
      stripHtml(htmlBody),
      mailOptions
    );

    // 🔹 Construction objet standardisé
    const campaign = buildCampaignData(
      "test",
      [me],
      subject,
      htmlBody,
      attachments || [],
      null
    );

    // 🔹 Enregistrement dans l’historique
    logCampaignHistory(
      campaign.name,
      campaign.recipients,
      new Date(),
      "Test envoyé",
      "Campagne de test directe",
      subject
    );

    Logger.log(`✅ Test envoyé à ${me}`);
    return `Test envoyé à ${me}`;
  } catch (e) {
    Logger.log("❌ Erreur sendTestToMe : " + e.message);
    throw e;
  }
}

/*****************************************************
 * ✅ Gestion des quotas améliorée : séparation vérif / incrément
 *****************************************************/

/**
 * Vérifie si l'utilisateur peut envoyer un nombre donné de destinataires.
 * NE MODIFIE PAS le quota.
 */
function checkQuota(numRecipients) {
  var config = getVersionConfig();
  var quota = getUserQuota();
  var now = new Date();

  // Reset mensuel si nouveau mois
  var last = new Date(quota.lastReset);
  if (
    now.getMonth() !== last.getMonth() ||
    now.getFullYear() !== last.getFullYear()
  ) {
    quota.monthly = 0;
    quota.lastReset = now.toISOString();
  }

  if (numRecipients > config.maxRecipients)
    throw new Error(
      `🚫 Limite par campagne dépassée : max ${config.maxRecipients} destinataires`
    );
  if (quota.monthly + numRecipients > config.monthlyQuota)
    throw new Error(
      `🚫 Limite mensuelle dépassée : max ${config.monthlyQuota} destinataires`
    );
  if (quota.annual + numRecipients > config.annualQuota)
    throw new Error(
      `🚫 Limite annuelle dépassée : max ${config.annualQuota} destinataires`
    );

  return true;
}

/**
 * Incrémente le quota utilisateur après un envoi réussi.
 */
function updateQuota(numRecipients) {
  var quota = getUserQuota();
  quota.monthly += numRecipients;
  quota.annual += numRecipients;
  quota.lastReset = quota.lastReset || new Date().toISOString();
  saveUserQuota(quota);
}

/*****************************************************
 * 🚀 Envoi d'une campagne (batch + quota incrémenté ici)
 *****************************************************/
function sendCampaign(recipients, subject, htmlBody, attachments) {
  if (!recipients || !recipients.length)
    throw new Error("Aucun destinataire fourni.");

  // ✅ Vérification des quotas avant envoi
  checkQuota(recipients.length);

  const campaign = buildCampaignData(
    "directe",
    recipients,
    subject,
    htmlBody,
    attachments
  );

  // Conversion pièces jointes
  const blobs = [];
  if (campaign.attachments?.length) {
    campaign.attachments.forEach((att) => {
      try {
        blobs.push(
          Utilities.newBlob(
            Utilities.base64Decode(att.base64),
            att.mime || "application/octet-stream",
            att.name
          )
        );
      } catch (e) {
        Logger.log("Erreur pièce jointe: " + att.name + " -> " + e.message);
      }
    });
  }

  // Envoi par batches
  const BATCH_SIZE_GMAIL = 40;
  let sent = 0;
  const errors = [];

  for (let i = 0; i < campaign.recipients.length; i += BATCH_SIZE_GMAIL) {
    const batch = campaign.recipients.slice(i, i + BATCH_SIZE_GMAIL);
    batch.forEach((r) => {
      try {
        const sub = replacePlaceholders(campaign.subject, r);
        const body = replacePlaceholders(campaign.htmlBody, r);

        GmailApp.sendEmail(r.email, sub, stripHtml(body), {
          htmlBody: body,
          name: SENDER_NAME,
          replyTo: PRO_EMAIL,
          attachments: blobs.length ? blobs : undefined,
        });

        sent++;
      } catch (e) {
        errors.push({ to: r.email, error: e.message });
      }
    });

    if (i + BATCH_SIZE_GMAIL < campaign.recipients.length)
      Utilities.sleep(1500);
  }

  // ✅ Mise à jour du quota après envoi réussi
  if (sent > 0) updateQuota(sent);

  logCampaignHistory(
    campaign.name,
    campaign.recipients,
    new Date(),
    errors.length ? "Partiel" : "Envoyé",
    errors.length ? "Erreurs (" + errors.length + ")" : "Succès",
    campaign.subject
  );

  return `Campagne directe terminée (${sent} envoyés, ${errors.length} erreurs)`;
}
/** ===============================
 * 5️⃣ stripHtml (helper)
 * Supprime les balises HTML pour le corps texte brut.
 * @param {string} html - Contenu HTML
 * @returns {string} Contenu texte brut
 */
function stripHtml(html) {
  return html.replace(/<[^>]+>/g, "");
}

/*****************************************************
 * NOVAMAIL — BLOC 3 FINAL HYBRIDE : PLANIFICATION, ENVOI & HISTORIQUE
 *****************************************************/

/**
 * 🔁 Remplace les variables {{nom}}, {{email}}, {{date}}
 */
function replacePlaceholders(text, recipient) {
  if (!text) return "";
  return text
    .replace(/{{\s*nom\s*}}/gi, recipient.nom || "")
    .replace(/{{\s*email\s*}}/gi, recipient.email || "")
    .replace(/{{\s*date\s*}}/gi, formatDateFR(new Date()));
}

/**
 * 🧹 Supprime les triggers d'une fonction donnée
 */
function removeTriggersForFunction(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  for (let t of triggers) {
    if (t.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(t);
      Logger.log("🗑️ Trigger supprimé : " + functionName);
    }
  }
}

/**
 * 🧠 Active un déclencheur maître unique (toutes les minutes)
 */
function ensureMasterTrigger() {
  removeTriggersForFunction("runAllScheduledCampaigns");
  ScriptApp.newTrigger("runAllScheduledCampaigns")
    .timeBased()
    .everyMinutes(1)
    .create();
  Logger.log("✅ Déclencheur maître actif : vérification toutes les minutes");
}

/**
 * 💾 Enregistre une campagne planifiée
 */

/**
 * 💾 Enregistre une campagne planifiée
 * ✅ Vérification des quotas à la planification
 */
/*****************************************************
 * 💾 Planification d'une campagne (ne consomme PAS le quota)
 *****************************************************/
function scheduleCampaign(campaignInput) {
  if (!campaignInput.recipients?.length)
    throw new Error("Aucun destinataire fourni.");
  if (!campaignInput.sendAt)
    throw new Error("Date de planification manquante.");

  const sendDate = new Date(campaignInput.sendAt);
  if (isNaN(sendDate)) throw new Error("Date invalide.");
  if (sendDate <= new Date())
    throw new Error("La date doit être dans le futur.");

  // ✅ Vérification des quotas uniquement pour validation (pas d'incrément)
  checkQuota(campaignInput.recipients.length);

  const campaign = buildCampaignData(
    "planifiée",
    campaignInput.recipients,
    campaignInput.subject,
    campaignInput.htmlBody,
    campaignInput.attachments,
    sendDate.toISOString()
  );

  const id = SCHEDULED_PREFIX + Utilities.getUuid();
  campaign.id = id;

  SCRIPT_PROPS.setProperty(id, JSON.stringify(campaign));

  Logger.log(
    `✅ Campagne planifiée : ${campaign.name} (${id}) pour ${formatDateFR(
      sendDate
    )}`
  );
  ensureMasterTrigger();

  return { success: true, id, scheduledAt: formatDateFR(sendDate) };
}

/**
 * 📊 Historique — sauvegarde chaque envoi dans la feuille
 */
/******************************************************
 * 📘 NOVAMAIL — GESTION HYBRIDE DE L’HISTORIQUE
 * ----------------------------------------------------
 * - Un seul fichier Google Sheet "NovaMail_Historique"
 * - Stocke toutes les campagnes envoyées (planifiées ou directes)
 * - Permet la consultation et la suppression via le front
 ******************************************************/

/**
 * 🗂️ Récupère ou crée le fichier Google Sheet d’historique unique.
 * Si le fichier n’existe pas, il est créé et son ID est stocké dans SCRIPT_PROPS.
 */
/******************************************************
 * 📘 NOVAMAIL — GESTION HYBRIDE DE L’HISTORIQUE (VERSION STABLE)
 * --------------------------------------------------------------
 * - Un seul fichier Google Sheet "NovaMail_Historique"
 * - ID stocké dans SCRIPT_PROPS pour éviter les doublons
 * - Chaque campagne envoyée (planifiée ou manuelle) est loguée
 ******************************************************/

/**
 * 🗂️ Récupère ou crée le fichier unique d’historique.
 * - Vérifie si un ID est stocké dans SCRIPT_PROPS
 * - Si le fichier n’existe plus, il est recréé une seule fois
 */
/******************************************************
 * 📘 NOVAMAIL — GESTION HYBRIDE DE L’HISTORIQUE (FIX 2025)
 * --------------------------------------------------------------
 * - Crée toujours un seul fichier visible dans ton Drive
 * - Stocke son ID dans SCRIPT_PROPS pour réutilisation
 * - Historique structuré et nettoyable
 ******************************************************/

/**
 * 🗂️ Récupère ou crée un fichier d’historique unique (visible dans Drive)
 */
/******************************************************
 * 📘 NOVAMAIL — GESTION HYBRIDE HISTORIQUE (FIX 2025-10-25)
 ******************************************************/

/******************************************************
 * 📘 NOVAMAIL — HISTORIQUE (VERSION FINALE STABLE)
 ******************************************************/

/**
 * 🧼 Réinitialise manuellement l’ID d’historique (debug)
 */
function resetHistoryFileId() {
  SCRIPT_PROPS.deleteProperty(HISTORY_PROP_KEY);
  Logger.log("🧽 ID d’historique supprimé → recréation au prochain appel.");
}

/**
 * 🗂️ Récupère ou crée le fichier d’historique unique
 */
function getOrCreateHistoryFile() {
  let fileId = SCRIPT_PROPS.getProperty(HISTORY_PROP_KEY);
  let file;

  // 1️⃣ Si un ID est enregistré → on teste s’il existe encore
  if (fileId) {
    try {
      file = DriveApp.getFileById(fileId);
      if (file && file.getId()) {
        Logger.log("📘 Fichier d’historique trouvé : " + fileId);
        return SpreadsheetApp.openById(fileId);
      }
    } catch (e) {
      Logger.log("⚠️ ID invalide → suppression et recréation...");
      SCRIPT_PROPS.deleteProperty(HISTORY_PROP_KEY);
    }
  }

  // 2️⃣ Sinon on cherche un fichier du même nom
  const existing = DriveApp.getFilesByName(HISTORY_FILE_NAME);
  if (existing.hasNext()) {
    file = existing.next();
    SCRIPT_PROPS.setProperty(HISTORY_PROP_KEY, file.getId());
    Logger.log("📄 Fichier existant réutilisé : " + file.getId());
    return SpreadsheetApp.openById(file.getId());
  }

  // 3️⃣ Sinon → on crée un nouveau fichier
  const ss = SpreadsheetApp.create(HISTORY_FILE_NAME);
  SCRIPT_PROPS.setProperty(HISTORY_PROP_KEY, ss.getId());
  Logger.log("🆕 Nouveau fichier d’historique créé : " + ss.getId());
  return ss;
}

/**
 * 📄 Récupère ou crée la feuille "Historique_Campagnes"
 */
function getHistorySheet() {
  const ss = getOrCreateHistoryFile();
  let sheet = ss.getSheetByName(HISTORY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(HISTORY_SHEET_NAME);
    sheet.appendRow([
      "Date d'envoi",
      "Nom de la campagne",
      "Objet",
      "Nombre de destinataires",
      "Emails",
      "Statut",
      "Détails / Erreur",
    ]);
  }
  return sheet;
}

/**
 * 📝 Ajoute une ligne d’historique
 */
function logCampaignHistory(
  campaignName,
  recipients,
  date,
  status = "Envoyé",
  details = "",
  subject = ""
) {
  try {
    const sheet = getHistorySheet();

    // ✅ Sécurise la date
    let sendDate = date instanceof Date ? date : new Date(date);
    if (isNaN(sendDate.getTime())) {
      sendDate = new Date();
      details += " (⚠️ Date invalide corrigée automatiquement)";
    }

    const nb = Array.isArray(recipients) ? recipients.length : 0;
    const emails = Array.isArray(recipients)
      ? recipients.map((r) => r.email || r).join(", ")
      : "";

    sheet.appendRow([
      Utilities.formatDate(
        sendDate,
        Session.getScriptTimeZone(),
        "dd/MM/yyyy HH:mm:ss"
      ),
      campaignName || "(Sans nom)",
      subject || "",
      nb,
      emails,
      status,
      details,
    ]);

    Logger.log(`📊 Historique ajouté : ${campaignName} (${nb} destinataires)`);
  } catch (e) {
    Logger.log("❌ Erreur logCampaignHistory : " + e.message);
  }
}

/**
 * 📤 Envoie l’historique complet au front-end
 */
function getHistoryData() {
  try {
    const sheet = getHistorySheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headers = data[0];
    const mapping = {
      "Date d'envoi": "date",
      "Nom de la campagne": "name",
      Objet: "subject",
      "Nombre de destinataires": "recipients",
      Statut: "status",
      "Détails / Erreur": "details",
    };

    return data.slice(1).map((r) => {
      const obj = {};
      headers.forEach((h, i) => {
        const key = mapping[h] || h;
        obj[key] = r[i];
      });
      return obj;
    });
  } catch (e) {
    Logger.log("❌ Erreur getHistoryData : " + e.message);
    return [];
  }
}

/**
 * 🧹 Vide l’historique (garde l’entête)
 */
function clearHistorySheet() {
  try {
    const sheet = getHistorySheet();
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
    Logger.log("🧹 Historique vidé.");
  } catch (e) {
    Logger.log("❌ Erreur clearHistorySheet : " + e.message);
  }
}

/**
 * 🔍 Debug : affiche l’ID actuel dans les logs
 */
function debugShowHistoryFileId() {
  const id = SCRIPT_PROPS.getProperty(HISTORY_PROP_KEY);
  Logger.log("📘 ID enregistré du fichier d’historique : " + id);
  return id;
}

/**
 * ⚙️ Déclencheur maître : envoie toutes les campagnes arrivées à échéance
 */

/**
 * 🔹 Déclencheur maître : envoie toutes les campagnes arrivées à échéance
 * ✅ Vérification quotas avant envoi
 */
function runAllScheduledCampaigns() {
  const props = SCRIPT_PROPS.getProperties();
  const now = new Date();
  Logger.log("⏰ Vérification campagnes planifiées à " + formatDateFR(now));

  let sentCount = 0;

  for (const key in props) {
    if (!key.startsWith(SCHEDULED_PREFIX)) continue;

    try {
      const campaign = JSON.parse(props[key]);
      const jobDate = new Date(campaign.sendAt || campaign.scheduledAt);

      if (now >= jobDate) {
        Logger.log(`🚀 Envoi ${campaign.name} (${formatDateFR(jobDate)})`);

        // ✅ Vérification quotas avant envoi
        try {
          checkQuota(campaign.recipients.length);
        } catch (quotaError) {
          logCampaignHistory(
            campaign.name,
            campaign.recipients,
            jobDate,
            "Erreur",
            quotaError.message
          );
          SCRIPT_PROPS.deleteProperty(key);
          continue;
        }

        // Pièces jointes
        const blobs = [];
        if (campaign.attachments?.length) {
          campaign.attachments.forEach((att) => {
            try {
              blobs.push(
                Utilities.newBlob(
                  Utilities.base64Decode(att.base64),
                  att.mime || "application/octet-stream",
                  att.name
                )
              );
            } catch (err) {
              Logger.log("Erreur pièce jointe planifiée: " + err.message);
            }
          });
        }

        // Envoi par batches
        const BATCH_SIZE_GMAIL = 40;
        for (let i = 0; i < campaign.recipients.length; i += BATCH_SIZE_GMAIL) {
          const batch = campaign.recipients.slice(i, i + BATCH_SIZE_GMAIL);
          batch.forEach((r) => {
            try {
              const sub = replacePlaceholders(campaign.subject, r);
              const body = replacePlaceholders(campaign.htmlBody, r);

              GmailApp.sendEmail(r.email, sub, stripHtml(body), {
                htmlBody: body,
                name: SENDER_NAME,
                replyTo: PRO_EMAIL,
                attachments: blobs.length ? blobs : undefined,
              });
            } catch (err) {
              logCampaignHistory(
                campaign.name,
                [r],
                jobDate,
                "Erreur",
                err.message
              );
            }
          });
          if (i + BATCH_SIZE_GMAIL < campaign.recipients.length)
            Utilities.sleep(1500);
        }

        logCampaignHistory(
          campaign.name,
          campaign.recipients,
          jobDate,
          "Envoyé",
          "Succès",
          campaign.subject
        );

        SCRIPT_PROPS.deleteProperty(key);
        sentCount++;
      }
    } catch (e) {
      Logger.log("Erreur planifiée: " + e.message);
    }
  }

  if (!sentCount) Logger.log("💤 Aucune campagne à exécuter.");
}

/**
 * 🔹 Liste toutes les campagnes planifiées (pour affichage côté front)
 * → Donne des champs cohérents avec la structure buildCampaignData()
 */
function getScheduledCampaigns() {
  const props = SCRIPT_PROPS.getProperties();
  const list = [];

  for (const key in props) {
    if (!key.startsWith(SCHEDULED_PREFIX)) continue;
    try {
      const job = JSON.parse(props[key]);

      list.push({
        id: job.id,
        name: job.name || "Campagne Planifiée",
        subject: job.subject || "(Sans objet)",
        htmlBody: job.htmlBody || "",
        sendAt: job.sendAt || job.scheduledAt || "",
        recipients: job.recipients || [],
        createdAt: job.createdAt || "",
      });
    } catch (e) {
      Logger.log("⚠️ Erreur parsing campagne " + key + ": " + e.message);
    }
  }

  // Tri des campagnes à venir par date croissante
  list.sort((a, b) => new Date(a.sendAt) - new Date(b.sendAt));
  Logger.log(`📋 ${list.length} campagne(s) planifiée(s) listée(s).`);
  return list;
}

/**
 * 🔹 Met à jour la date planifiée d’une campagne existante
 * → Conserve toutes les autres propriétés intactes
 */
function updateScheduledDate(id, newSendAtIso) {
  if (!id || !newSendAtIso) throw new Error("⛔ id et nouvelle date requis.");

  const prop = SCRIPT_PROPS.getProperty(id);
  if (!prop) throw new Error("Campagne introuvable: " + id);

  const job = JSON.parse(prop);
  const newDate = new Date(newSendAtIso);
  if (isNaN(newDate)) throw new Error("Date invalide : " + newSendAtIso);

  job.sendAt = newDate.toISOString();
  SCRIPT_PROPS.setProperty(id, JSON.stringify(job));

  Logger.log("🕒 Campagne replanifiée : " + id + " → " + formatDateFR(newDate));
  return { success: true, id, scheduledAt: formatDateFR(newDate) };
}

/**
 * 🔹 Supprime une campagne planifiée avant exécution
 * → Nettoie proprement l’entrée dans SCRIPT_PROPS
 */
function deleteScheduledCampaign(id) {
  if (!id) throw new Error("id requis pour suppression.");

  const prop = SCRIPT_PROPS.getProperty(id);
  if (!prop) {
    Logger.log("⚠️ Campagne introuvable : " + id);
    return { success: false, message: "Campagne introuvable." };
  }

  SCRIPT_PROPS.deleteProperty(id);
  Logger.log("🗑️ Campagne supprimée : " + id);
  return { success: true, id };
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index").setTitle(
    "NovaMail F.J.C"
  );
}
