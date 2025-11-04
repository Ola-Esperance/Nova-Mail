/*****************************************************
 * NOVAMAIL SAAS - CAMPAIGNS.GS
 * ====================================================
 * Gestion complète des campagnes email
 * Envoi direct, gestion des destinataires, templates, pièces jointes
 * 
 * @author NovaMail Team
 * @version 2.0.0
 * @lastModified 2025-11-03
 *****************************************************/

/**
 * ============================================
 * STRUCTURE DE DONNÉES CAMPAGNE
 * ============================================
 */

/**
 * Définition de l'objet Campaign
 * @typedef {Object} Campaign
 * @property {string} id - Identifiant unique campagne
 * @property {string} type - Type: "direct" | "scheduled" | "test"
 * @property {string} name - Nom de la campagne
 * @property {string} subject - Objet de l'email
 * @property {string} htmlBody - Contenu HTML
 * @property {Array<Object>} recipients - Destinataires [{nom, email}]
 * @property {Array<Object>} attachments - Pièces jointes [{name, base64, mime}]
 * @property {string} senderEmail - Email expéditeur (optionnel)
 * @property {string} senderName - Nom expéditeur
 * @property {string} replyTo - Email de réponse
 * @property {string} createdAt - Date création (ISO)
 * @property {string} sentAt - Date envoi (ISO, si envoyé)
 * @property {string} status - Statut: "draft" | "sending" | "sent" | "failed"
 * @property {Object} stats - Statistiques {sent, failed, errors}
 */

/**
 * ============================================
 * IMPORT & GESTION DES DESTINATAIRES
 * ============================================
 */

/**
 * 📋 Liste tous les Google Sheets accessibles à l'utilisateur
 * Utilisé pour le sélecteur d'import dans le frontend
 * 
 * @returns {Array<Object>} Liste des sheets [{id, name, modifiedDate}]
 * 
 * @example
 * const sheets = listGoogleSheets();
 * // [{id: "1abc...", name: "Contacts", modifiedDate: "2025-11-03T10:00:00Z"}]
 */
function listGoogleSheets() {
  try {
    const files = DriveApp.searchFiles(
      "mimeType='application/vnd.google-apps.spreadsheet'"
    );
    
    const sheetsList = [];
    
    while (files.hasNext()) {
      const file = files.next();
      sheetsList.push({
        id: file.getId(),
        name: file.getName(),
        modifiedDate: file.getLastUpdated() 
          ? file.getLastUpdated().toISOString() 
          : ""
      });
    }
    
    // Tri par date de modification décroissante
    sheetsList.sort((a, b) => b.modifiedDate.localeCompare(a.modifiedDate));
    
    return sheetsList;
    
  } catch (error) {
    logError("listGoogleSheets", error);
    throw new Error("Impossible de lister les Google Sheets : " + error.message);
  }
}

/**
 * 📊 Importe des destinataires depuis un Google Sheet
 * Détecte automatiquement les colonnes "nom" et "email"
 * 
 * @param {string} fileId - ID du Google Sheet
 * @param {string} sheetName - Nom de la feuille (optionnel)
 * @returns {Array<Object>} Destinataires [{nom, email}]
 * 
 * @example
 * const recipients = importFromSheet("1abc...XYZ", "Contacts");
 */
function importFromSheet(fileId, sheetName) {
  if (!fileId) {
    throw new Error("ID du fichier manquant");
  }
  
  try {
    // Ouverture du spreadsheet
    const spreadsheet = SpreadsheetApp.openById(fileId);
    
    // Sélection de la feuille
    let sheet;
    if (sheetName) {
      sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        throw new Error(`Feuille "${sheetName}" introuvable`);
      }
    } else {
      sheet = spreadsheet.getSheets()[0];
    }
    
    if (!sheet) {
      throw new Error("Aucune feuille disponible dans ce fichier");
    }
    
    // Lecture des données
    const data = sheet.getDataRange().getValues();
    
    if (!data || data.length < 2) {
      throw new Error("Le fichier est vide ou ne contient pas d'en-têtes");
    }
    
    // Extraction des en-têtes (ligne 1)
    const headers = data[0].map(h => String(h).toLowerCase().trim());
    
    // Détection automatique des colonnes
    let nameIndex = -1;
    let emailIndex = -1;
    
    headers.forEach((header, index) => {
      if (header.includes("nom") || header.includes("name") || header.includes("prenom")) {
        nameIndex = index;
      }
      if (header.includes("mail") || header.includes("email") || header.includes("courriel")) {
        emailIndex = index;
      }
    });
    
    if (emailIndex === -1) {
      throw new Error(
        "Colonne 'email' introuvable. Assurez-vous qu'une colonne contient 'email' ou 'mail' dans son nom."
      );
    }
    
    // Extraction des destinataires
    const recipients = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const email = normalizeEmail(row[emailIndex]);
      
      // Validation email
      if (!isValidEmail(email)) continue;
      
      // Extraction du nom
      let name = "";
      if (nameIndex !== -1 && row[nameIndex]) {
        name = String(row[nameIndex]).trim();
      }
      
      // Génération nom si absent
      if (!name) {
        name = generateNameFromEmail(email);
      }
      
      recipients.push({ nom: name, email: email });
    }
    
    // Dédoublonnage par email
    const uniqueRecipients = [];
    const seenEmails = new Set();
    
    recipients.forEach(r => {
      if (!seenEmails.has(r.email)) {
        seenEmails.add(r.email);
        uniqueRecipients.push(r);
      }
    });
    
    // Sauvegarde comme dernière liste importée
    saveLastImportedList(uniqueRecipients);
    
    logInfo(`✅ Import réussi : ${uniqueRecipients.length} destinataires depuis "${sheet.getName()}"`);
    
    return uniqueRecipients;
    
  } catch (error) {
    logError("importFromSheet", error);
    throw new Error("Erreur lors de l'import : " + error.message);
  }
}

/**
 * 📄 Importe des destinataires depuis un fichier CSV
 * Accepte CSV brut ou base64
 * 
 * @param {string} csvContent - Contenu CSV (texte brut ou base64)
 * @returns {Array<Object>} Destinataires [{nom, email}]
 * 
 * @example
 * const csv = "Nom,Email\nJohn Doe,john@example.com";
 * const recipients = importCsv(csv);
 */
function importCsv(csvContent) {
  if (!csvContent) {
    throw new Error("Contenu CSV vide");
  }
  
  try {
    // Détection si base64
    let csvText = csvContent;
    if (csvContent.match(/^[A-Za-z0-9+/=]+$/)) {
      try {
        csvText = Utilities.newBlob(
          Utilities.base64Decode(csvContent)
        ).getDataAsString();
      } catch (e) {
        // Pas du base64, considérer comme texte brut
      }
    }
    
    // Conversion en destinataires
    const recipients = convertCsvToRecipients(csvText);
    
    // Sauvegarde dernière liste
    saveLastImportedList(recipients);
    
    logInfo(`✅ Import CSV réussi : ${recipients.length} destinataires`);
    
    return recipients;
    
  } catch (error) {
    logError("importCsv", error);
    throw new Error("Erreur lors de l'import CSV : " + error.message);
  }
}

/**
 * Convertit un texte CSV en tableau de destinataires
 * Supporte : "Nom,Email" ou "Email" seul
 * 
 * @param {string} csvText - Texte CSV
 * @returns {Array<Object>} Destinataires
 */
function convertCsvToRecipients(csvText) {
  if (!csvText) return [];
  
  const lines = csvText
    .split(/\r?\n/)
    .map(line => line.trim())
    .filter(Boolean);
  
  const recipients = [];
  
  lines.forEach((line, index) => {
    // Ignorer la ligne d'en-tête si elle contient "nom" ou "email"
    if (index === 0 && line.toLowerCase().includes("email")) {
      return;
    }
    
    // Séparation par virgule (gestion guillemets basique)
    const parts = line.split(",").map(p => p.trim().replace(/^["']|["']$/g, ""));
    
    let name = "";
    let email = "";
    
    if (parts.length >= 2) {
      name = parts[0];
      email = parts[1];
    } else {
      email = parts[0];
    }
    
    // Validation et ajout
    email = normalizeEmail(email);
    if (!isValidEmail(email)) return;
    
    if (!name) {
      name = generateNameFromEmail(email);
    }
    
    recipients.push({ nom: name, email: email });
  });
  
  return recipients;
}

/**
 * ⤴️ Récupère la dernière liste de destinataires importée
 * 
 * @returns {Array<Object>} Dernière liste ou []
 */
function getLastImportedList() {
  return loadUserProperty(USER_STORAGE_KEYS.LAST_LIST) || [];
}

/**
 * 💾 Sauvegarde la dernière liste importée
 * 
 * @param {Array<Object>} recipients - Liste destinataires
 * @returns {boolean} Succès
 */
function saveLastImportedList(recipients) {
  if (!recipients || !recipients.length) return false;
  return saveUserProperty(USER_STORAGE_KEYS.LAST_LIST, recipients);
}

/**
 * 🧹 Efface la dernière liste importée
 * 
 * @returns {boolean} Succès
 */
function clearLastImportedList() {
  return deleteUserProperty(USER_STORAGE_KEYS.LAST_LIST);
}

/**
 * ============================================
 * GESTION DES TEMPLATES
 * ============================================
 */

/**
 * 💾 Sauvegarde un template d'email
 * 
 * @param {string} name - Nom du template (unique par utilisateur)
 * @param {string} subject - Sujet
 * @param {string} htmlBody - Contenu HTML
 * @returns {Object} Template sauvegardé
 * 
 * @example
 * saveTemplate("Promo Novembre", "🎉 Offre spéciale", "<p>Bonjour {{nom}}...</p>");
 */
function saveTemplate(name, subject, htmlBody) {
  if (!name || !name.trim()) {
    throw new Error("Le nom du template est obligatoire");
  }
  
  // Vérification quota templates selon version
  const config = getVersionConfig();
  if (config.maxTemplates === 0) {
    throw new Error(
      `❌ Sauvegarde de templates non disponible avec votre version ${config.name}. ` +
      `Passez à STARTER ou supérieur.`
    );
  }
  
  // Vérification nombre de templates existants
  if (config.maxTemplates > 0) {
    const existing = listTemplates();
    if (existing.length >= config.maxTemplates) {
      throw new Error(
        `❌ Limite de templates atteinte (${config.maxTemplates}). ` +
        `Supprimez un template existant ou passez à une version supérieure.`
      );
    }
  }
  
  const template = {
    name: name.trim(),
    subject: subject || "",
    htmlBody: htmlBody || "",
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };
  
  // Sauvegarde avec préfixe TEMPLATE:
  const key = STORAGE_PREFIX.USER_PROPS + ":" + STORAGE_PREFIX.TEMPLATE + name.trim();
  PropertiesService.getUserProperties().setProperty(key, JSON.stringify(template));
  
  logInfo(`💾 Template sauvegardé : "${name}"`);
  
  return template;
}

/**
 * 📄 Charge un template par nom
 * 
 * @param {string} name - Nom du template
 * @returns {Object|null} Template ou null
 */
function loadTemplate(name) {
  if (!name) {
    throw new Error("Nom du template requis");
  }
  
  try {
    const key = STORAGE_PREFIX.USER_PROPS + ":" + STORAGE_PREFIX.TEMPLATE + name.trim();
    const raw = PropertiesService.getUserProperties().getProperty(key);
    
    if (!raw) return null;
    
    return JSON.parse(raw);
    
  } catch (error) {
    logError("loadTemplate", error);
    return null;
  }
}

/**
 * 📚 Liste tous les templates de l'utilisateur
 * 
 * @returns {Array<Object>} Liste des templates triés par date
 */
function listTemplates() {
  try {
    const props = PropertiesService.getUserProperties().getProperties();
    const templates = [];
    const prefix = STORAGE_PREFIX.USER_PROPS + ":" + STORAGE_PREFIX.TEMPLATE;
    
    for (const key in props) {
      if (key.startsWith(prefix)) {
        try {
          const template = JSON.parse(props[key]);
          templates.push(template);
        } catch (e) {
          // Ignorer les entrées corrompues
          logWarning(`Template corrompu ignoré : ${key}`);
        }
      }
    }
    
    // Tri par date de mise à jour décroissante
    templates.sort((a, b) => 
      (b.updatedAt || "").localeCompare(a.updatedAt || "")
    );
    
    return templates;
    
  } catch (error) {
    logError("listTemplates", error);
    return [];
  }
}

/**
 * 🗑️ Supprime un template
 * 
 * @param {string} name - Nom du template
 * @returns {boolean} Succès
 */
function deleteTemplate(name) {
  if (!name) {
    throw new Error("Nom du template requis");
  }
  
  try {
    const key = STORAGE_PREFIX.USER_PROPS + ":" + STORAGE_PREFIX.TEMPLATE + name.trim();
    const props = PropertiesService.getUserProperties();
    
    if (!props.getProperty(key)) {
      throw new Error(`Template "${name}" introuvable`);
    }
    
    props.deleteProperty(key);
    logInfo(`🗑️ Template supprimé : "${name}"`);
    
    return true;
    
  } catch (error) {
    logError("deleteTemplate", error);
    throw error;
  }
}

/**
 * 🧹 Supprime tous les templates de l'utilisateur
 * 
 * @returns {number} Nombre de templates supprimés
 */
function deleteAllTemplates() {
  try {
    const props = PropertiesService.getUserProperties();
    const allProps = props.getProperties();
    const prefix = STORAGE_PREFIX.USER_PROPS + ":" + STORAGE_PREFIX.TEMPLATE;
    
    let count = 0;
    
    for (const key in allProps) {
      if (key.startsWith(prefix)) {
        props.deleteProperty(key);
        count++;
      }
    }
    
    logInfo(`🧹 ${count} template(s) supprimé(s)`);
    
    return count;
    
  } catch (error) {
    logError("deleteAllTemplates", error);
    return 0;
  }
}

/**
 * ============================================
 * PRÉVISUALISATION
 * ============================================
 */

/**
 * 🔍 Génère une prévisualisation d'email avec placeholders remplacés
 * 
 * @param {string} subject - Sujet de l'email
 * @param {string} htmlBody - Contenu HTML
 * @param {Object} recipient - Destinataire exemple {nom, email}
 * @returns {Object} Prévisualisation {subject, htmlBody}
 * 
 * @example
 * const preview = previewEmail("Bonjour {{nom}}", "<p>Email pour {{email}}</p>", 
 *   {nom: "John", email: "john@example.com"}
 * );
 */
function previewEmail(subject, htmlBody, recipient) {
  if (!recipient) {
    recipient = { nom: "Utilisateur Exemple", email: "exemple@email.com" };
  }
  
  return {
    subject: replacePlaceholders(subject, recipient),
    htmlBody: replacePlaceholders(htmlBody, recipient)
  };
}

/**
 * ============================================
 * ENVOI DE CAMPAGNES
 * ============================================
 */

/**
 * 📤 Envoie une campagne directe (immédiate)
 * Gère automatiquement :
 * - Vérification quotas
 * - Envoi par batches
 * - Remplacement placeholders
 * - Pièces jointes
 * - Logging historique
 * 
 * @param {Array<Object>} recipients - Destinataires [{nom, email}]
 * @param {string} subject - Sujet
 * @param {string} htmlBody - Contenu HTML
 * @param {Array<Object>} attachments - Pièces jointes (optionnel)
 * @returns {Object} Résultat {success, sent, failed, errors}
 * 
 * @example
 * const result = sendCampaign(
 *   [{nom: "John", email: "john@example.com"}],
 *   "Newsletter Novembre",
 *   "<p>Bonjour {{nom}},...</p>",
 *   []
 * );
 */
function sendCampaign(recipients, subject, htmlBody, attachments) {
  try {
    // ✅ Validation des paramètres
    if (!recipients || !Array.isArray(recipients) || recipients.length === 0) {
      throw new Error("Aucun destinataire fourni");
    }
    
    if (!subject || !subject.trim()) {
      throw new Error("Objet de l'email requis");
    }
    
    if (!htmlBody || !htmlBody.trim()) {
      throw new Error("Contenu de l'email requis");
    }
    
    // ✅ Vérification des quotas (ne consomme PAS encore)
    checkQuotaAvailable(recipients.length);
    
    // ✅ Récupération configuration utilisateur
    const config = getVersionConfig();
    const client = loadClientByCurrentUser();
    
    // Construction objet campagne
    const campaign = {
      id: generateShortId(),
      type: "direct",
      name: subject || "Campagne directe",
      subject: subject,
      htmlBody: htmlBody,
      recipients: recipients,
      attachments: attachments || [],
      senderEmail: client ? client.senderEmail : DEFAULT_SENDER_EMAIL,
      senderName: client ? client.companyName || client.fullName : DEFAULT_SENDER_NAME,
      replyTo: client ? client.replyEmail : DEFAULT_SENDER_EMAIL,
      createdAt: new Date().toISOString(),
      status: "sending"
    };
    
    // ✅ Conversion des pièces jointes
    const blobs = [];
    
    if (campaign.attachments && campaign.attachments.length > 0) {
      // Vérification si pièces jointes autorisées
      if (!config.allowAttachments) {
        throw new Error(
          `❌ Pièces jointes non disponibles avec votre version ${config.name}. ` +
          `Passez à STARTER ou supérieur.`
        );
      }
      
      campaign.attachments.forEach(att => {
        try {
          // Validation taille
          validateFileSize(att.base64, MAX_ATTACHMENT_SIZE_MB);
          
          // Conversion base64 → Blob
          const blob = base64ToBlob(att.base64, att.name, att.mime);
          blobs.push(blob);
          
        } catch (err) {
          logError("Conversion pièce jointe", err);
          throw new Error(`Erreur avec la pièce jointe "${att.name}" : ${err.message}`);
        }
      });
    }
    
    // ✅ Envoi par batches
    let sentCount = 0;
    const errors = [];
    
    for (let i = 0; i < recipients.length; i += GMAIL_BATCH_SIZE) {
      const batch = recipients.slice(i, i + GMAIL_BATCH_SIZE);
      
      batch.forEach(recipient => {
        try {
          // Remplacement des placeholders
          const personalizedSubject = replacePlaceholders(campaign.subject, recipient);
          const personalizedBody = replacePlaceholders(campaign.htmlBody, recipient);
          
          // Options d'envoi Gmail
          const mailOptions = {
            htmlBody: personalizedBody,
            name: campaign.senderName,
            replyTo: campaign.replyTo
          };
          
          // Ajout pièces jointes si présentes
          if (blobs.length > 0) {
            mailOptions.attachments = blobs;
          }
          
          // Envoi via Gmail
          GmailApp.sendEmail(
            recipient.email,
            personalizedSubject,
            stripHtml(personalizedBody), // Version texte brut
            mailOptions
          );
          
          sentCount++;
          
        } catch (err) {
          errors.push({
            email: recipient.email,
            error: err.message
          });
          logError(`Envoi à ${recipient.email}`, err);
        }
      });
      
      // Délai entre batches (évite throttling Gmail)
      if (i + GMAIL_BATCH_SIZE < recipients.length) {
        Utilities.sleep(GMAIL_BATCH_DELAY_MS);
      }
    }
    
    // ✅ Mise à jour du statut
    campaign.status = errors.length === 0 ? "sent" : "partial";
    campaign.sentAt = new Date().toISOString();
    campaign.stats = {
      sent: sentCount,
      failed: errors.length,
      errors: errors
    };
    
    // ✅ Incrémentation des quotas (seulement pour envois réussis)
    if (sentCount > 0) {
      incrementQuota(sentCount);
    }
    
    // ✅ Enregistrement dans l'historique
    logCampaignHistory(
      campaign.name,
      recipients,
      new Date(),
      campaign.status === "sent" ? "Envoyé" : "Partiel",
      errors.length > 0 ? `${errors.length} erreur(s)` : "Succès",
      campaign.subject
    );
    
    logInfo(
      `✅ Campagne "${campaign.name}" terminée : ` +
      `${sentCount} envoyés, ${errors.length} erreurs`
    );
    
    return {
      success: true,
      sent: sentCount,
      failed: errors.length,
      errors: errors,
      message: `Campagne envoyée : ${sentCount}/${recipients.length} destinataires atteints`
    };
    
  } catch (error) {
    logError("sendCampaign", error);
    throw error;
  }
}

/**
 * ✉️ Envoie un email de test à l'utilisateur courant
 * 
 * @param {string} subject - Sujet
 * @param {string} htmlBody - Contenu HTML
 * @param {Array<Object>} attachments - Pièces jointes (optionnel)
 * @returns {string} Message de confirmation
 * 
 * @example
 * sendTestToMe("Test", "<p>Contenu test</p>", []);
 */
function sendTestToMe(subject, htmlBody, attachments) {
  try {
    // Détermination de l'email destinataire
    const myEmail = Session.getActiveUser().getEmail() || 
                    Session.getEffectiveUser().getEmail();
    
    if (!myEmail || !isValidEmail(myEmail)) {
      throw new Error("Impossible de déterminer votre adresse email");
    }
    
    // Conversion pièces jointes
    const blobs = [];
    if (attachments && attachments.length > 0) {
      attachments.forEach(att => {
        try {
          blobs.push(base64ToBlob(att.base64, att.name, att.mime));
        } catch (err) {
          logWarning(`Pièce jointe ignorée : ${att.name}`);
        }
      });
    }
    
    // Options d'envoi
    const mailOptions = {
      htmlBody: htmlBody,
      name: DEFAULT_SENDER_NAME + " (Test)",
      replyTo: DEFAULT_SENDER_EMAIL
    };
    
    if (blobs.length > 0) {
      mailOptions.attachments = blobs;
    }
    
    // Envoi
    GmailApp.sendEmail(
      myEmail,
      "[TEST] " + subject,
      stripHtml(htmlBody),
      mailOptions
    );
    
    // Log dans l'historique
    logCampaignHistory(
      "Test",
      [myEmail],
      new Date(),
      "Envoyé",
      "Email de test",
      subject
    );
    
    logInfo(`✅ Test envoyé à ${myEmail}`);
    
    return `Email de test envoyé avec succès à ${myEmail}`;
    
  } catch (error) {
    logError("sendTestToMe", error);
    throw error;
  }
}

/**
 * ============================================
 * HELPERS INTERNES
 * ============================================
 */

/**
 * Charge les informations du client courant (si connecté via userId)
 * 
 * @returns {Client|null} Client ou null
 */
function loadClientByCurrentUser() {
  try {
    // Tentative de récupération depuis la session
    const userEmail = Session.getActiveUser().getEmail();
    
    if (userEmail && isValidEmail(userEmail)) {
      return findClientByEmail(userEmail);
    }
    
    return null;
    
  } catch (error) {
    return null;
  }
}
