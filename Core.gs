/*****************************************************
 * NOVAMAIL SAAS - CORE.GS
 * ====================================================
 * Fonctions utilitaires réutilisables dans tout le système
 * 
 * @author NovaMail Team
 * @version 2.0.0
 * @lastModified 2025-11-03
 *****************************************************/

/**
 * ============================================
 * VALIDATION & NORMALISATION EMAILS
 * ============================================
 */

/**
 * Normalise une adresse email (trim + lowercase)
 * 
 * @param {string} email - Email brut à normaliser
 * @returns {string} Email normalisé
 * 
 * @example
 * normalizeEmail(" John@EXAMPLE.com ") // "john@example.com"
 */
function normalizeEmail(email) {
  if (!email) return "";
  return String(email).trim().toLowerCase();
}

/**
 * Valide le format d'une adresse email
 * 
 * @param {string} email - Email à valider
 * @returns {boolean} true si format valide
 * 
 * @example
 * isValidEmail("test@example.com") // true
 * isValidEmail("invalid.email") // false
 */
function isValidEmail(email) {
  // Regex RFC 5322 simplifiée
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
  return emailRegex.test(normalizeEmail(email));
}

/**
 * Valide un tableau d'emails et retourne uniquement les valides
 * 
 * @param {Array<string>} emails - Tableau d'emails
 * @returns {Array<string>} Emails valides uniquement
 */
function filterValidEmails(emails) {
  if (!Array.isArray(emails)) return [];
  
  return emails
    .map(normalizeEmail)
    .filter(isValidEmail)
    .filter((email, index, self) => self.indexOf(email) === index); // Dédoublonnage
}

/**
 * ============================================
 * PARSING DE DESTINATAIRES
 * ============================================
 */

/**
 * Parse une liste de destinataires depuis du texte brut
 * Accepte les formats :
 * - "Nom,Email" (CSV-like)
 * - "Email" seul
 * 
 * @param {string} text - Texte multiligne contenant les destinataires
 * @returns {Array<Object>} Tableau d'objets {nom, email}
 * 
 * @example
 * parseRecipientsFromText("John Doe,john@example.com\njane@example.com")
 * // [
 * //   {nom: "John Doe", email: "john@example.com"},
 * //   {nom: "Jane", email: "jane@example.com"}
 * // ]
 */
function parseRecipientsFromText(text) {
  if (!text) return [];
  
  const lines = String(text)
    .split(/\r?\n/)
    .map(line => line.trim())
    .filter(Boolean);
  
  const recipients = [];
  
  lines.forEach(line => {
    const parts = line.split(",").map(p => p.trim());
    let name = "";
    let email = "";
    
    if (parts.length >= 2) {
      // Format "Nom,Email"
      name = parts[0];
      email = parts[1];
    } else {
      // Format "Email" seul
      email = parts[0];
    }
    
    // Normalisation et validation
    email = normalizeEmail(email);
    if (!isValidEmail(email)) return;
    
    // Génération automatique du nom si absent
    if (!name) {
      name = generateNameFromEmail(email);
    }
    
    recipients.push({ nom: name, email: email });
  });
  
  return recipients;
}

/**
 * Génère un nom lisible depuis une adresse email
 * 
 * @param {string} email - Adresse email
 * @returns {string} Nom généré
 * 
 * @example
 * generateNameFromEmail("john.doe@example.com") // "John Doe"
 */
function generateNameFromEmail(email) {
  const localPart = email.split("@")[0];
  
  return localPart
    .replace(/[._-]+/g, " ")              // Remplace séparateurs par espaces
    .replace(/\b\w/g, c => c.toUpperCase()); // Capitalize première lettre
}

/**
 * ============================================
 * GESTION DU STORAGE (PROPRIÉTÉS)
 * ============================================
 */

/**
 * Sauvegarde une valeur dans les propriétés utilisateur
 * 
 * @param {string} key - Clé de stockage
 * @param {*} value - Valeur à stocker (sera sérialisée en JSON)
 * @returns {boolean} Succès de l'opération
 */
function saveUserProperty(key, value) {
  if (!key) return false;
  
  try {
    const fullKey = STORAGE_PREFIX.USER_PROPS + ":" + key;
    PropertiesService.getUserProperties()
      .setProperty(fullKey, JSON.stringify(value));
    return true;
  } catch (error) {
    logError("saveUserProperty", error);
    return false;
  }
}

/**
 * Charge une valeur depuis les propriétés utilisateur
 * 
 * @param {string} key - Clé de stockage
 * @returns {*} Valeur désérialisée ou null si absente
 */
function loadUserProperty(key) {
  if (!key) return null;
  
  try {
    const fullKey = STORAGE_PREFIX.USER_PROPS + ":" + key;
    const raw = PropertiesService.getUserProperties().getProperty(fullKey);
    
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (error) {
    logError("loadUserProperty", error);
    return null;
  }
}

/**
 * Supprime une propriété utilisateur
 * 
 * @param {string} key - Clé à supprimer
 * @returns {boolean} Succès de l'opération
 */
function deleteUserProperty(key) {
  if (!key) return false;
  
  try {
    const fullKey = STORAGE_PREFIX.USER_PROPS + ":" + key;
    PropertiesService.getUserProperties().deleteProperty(fullKey);
    return true;
  } catch (error) {
    logError("deleteUserProperty", error);
    return false;
  }
}

/**
 * Sauvegarde une propriété script (globale)
 * 
 * @param {string} key - Clé de stockage
 * @param {*} value - Valeur à stocker
 * @returns {boolean} Succès de l'opération
 */
function saveScriptProperty(key, value) {
  if (!key) return false;
  
  try {
    const fullKey = STORAGE_PREFIX.SCRIPT_PROPS + ":" + key;
    PropertiesService.getScriptProperties()
      .setProperty(fullKey, JSON.stringify(value));
    return true;
  } catch (error) {
    logError("saveScriptProperty", error);
    return false;
  }
}

/**
 * Charge une propriété script (globale)
 * 
 * @param {string} key - Clé de stockage
 * @returns {*} Valeur désérialisée ou null
 */
function loadScriptProperty(key) {
  if (!key) return null;
  
  try {
    const fullKey = STORAGE_PREFIX.SCRIPT_PROPS + ":" + key;
    const raw = PropertiesService.getScriptProperties().getProperty(fullKey);
    
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (error) {
    logError("loadScriptProperty", error);
    return null;
  }
}

/**
 * Supprime une propriété script
 * 
 * @param {string} key - Clé à supprimer
 * @returns {boolean} Succès de l'opération
 */
function deleteScriptProperty(key) {
  if (!key) return false;
  
  try {
    const fullKey = STORAGE_PREFIX.SCRIPT_PROPS + ":" + key;
    PropertiesService.getScriptProperties().deleteProperty(fullKey);
    return true;
  } catch (error) {
    logError("deleteScriptProperty", error);
    return false;
  }
}

/**
 * Liste toutes les propriétés script avec le préfixe NovaMail
 * 
 * @returns {Object} Dictionnaire clé -> valeur
 */
function listAllScriptProperties() {
  try {
    const allProps = PropertiesService.getScriptProperties().getProperties();
    const filtered = {};
    const prefix = STORAGE_PREFIX.SCRIPT_PROPS + ":";
    
    for (const key in allProps) {
      if (key.startsWith(prefix)) {
        const cleanKey = key.substring(prefix.length);
        try {
          filtered[cleanKey] = JSON.parse(allProps[key]);
        } catch (e) {
          filtered[cleanKey] = allProps[key]; // Valeur brute si pas JSON
        }
      }
    }
    
    return filtered;
  } catch (error) {
    logError("listAllScriptProperties", error);
    return {};
  }
}

/**
 * ============================================
 * FORMATAGE DE DATES
 * ============================================
 */

/**
 * Formate une date en français lisible
 * 
 * @param {Date|string} date - Date à formater
 * @returns {string} Date formatée (ex: "03 novembre 2025 à 14:30:15")
 * 
 * @example
 * formatDateFR(new Date()) // "03 novembre 2025 à 14:30:15"
 */
function formatDateFR(date) {
  const d = date instanceof Date ? date : new Date(date);
  
  if (isNaN(d.getTime())) {
    return "(Date invalide)";
  }
  
  const options = {
    day: "2-digit",
    month: "long",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit"
  };
  
  return d.toLocaleString("fr-FR", options);
}

/**
 * Formate une date pour affichage court
 * 
 * @param {Date|string} date - Date à formater
 * @returns {string} Date formatée (ex: "03/11/2025 14:30")
 */
function formatDateShort(date) {
  const d = date instanceof Date ? date : new Date(date);
  
  if (isNaN(d.getTime())) {
    return "(Date invalide)";
  }
  
  return Utilities.formatDate(
    d,
    Session.getScriptTimeZone(),
    "dd/MM/yyyy HH:mm"
  );
}

/**
 * ============================================
 * MANIPULATION DE TEXTE
 * ============================================
 */

/**
 * Supprime les balises HTML d'une chaîne
 * Utile pour la version texte brut des emails
 * 
 * @param {string} html - Contenu HTML
 * @returns {string} Texte brut sans balises
 * 
 * @example
 * stripHtml("<p>Bonjour <strong>monde</strong></p>") // "Bonjour monde"
 */
function stripHtml(html) {
  if (!html) return "";
  return String(html)
    .replace(/<[^>]+>/g, "")          // Supprime balises
    .replace(/&nbsp;/g, " ")           // Remplace &nbsp;
    .replace(/&[a-z]+;/gi, "")        // Supprime entités HTML
    .trim();
}

/**
 * Remplace les placeholders dans un template
 * Supporte : {{nom}}, {{email}}, {{date}}
 * 
 * @param {string} text - Texte avec placeholders
 * @param {Object} recipient - Objet destinataire {nom, email}
 * @returns {string} Texte avec placeholders remplacés
 * 
 * @example
 * replacePlaceholders("Bonjour {{nom}}", {nom: "John", email: "j@ex.com"})
 * // "Bonjour John"
 */
function replacePlaceholders(text, recipient) {
  if (!text) return "";
  if (!recipient) return text;
  
  return text
    .replace(/{{\s*nom\s*}}/gi, recipient.nom || recipient.name || "")
    .replace(/{{\s*email\s*}}/gi, recipient.email || "")
    .replace(/{{\s*date\s*}}/gi, formatDateFR(new Date()));
}

/**
 * ============================================
 * GÉNÉRATION D'IDENTIFIANTS UNIQUES
 * ============================================
 */

/**
 * Génère un identifiant unique court (8 caractères)
 * 
 * @returns {string} ID unique
 * 
 * @example
 * generateShortId() // "a3f8k2p1"
 */
function generateShortId() {
  return Utilities.getUuid().substring(0, 8);
}

/**
 * Génère un identifiant unique complet (UUID v4)
 * 
 * @returns {string} UUID complet
 */
function generateFullId() {
  return Utilities.getUuid();
}

/**
 * ============================================
 * GESTION DES FICHIERS BASE64
 * ============================================
 */

/**
 * Convertit un fichier base64 en Blob Google Drive
 * 
 * @param {string} base64String - Données encodées en base64
 * @param {string} filename - Nom du fichier
 * @param {string} mimeType - Type MIME du fichier
 * @returns {Blob} Blob Google Apps Script
 * 
 * @example
 * const blob = base64ToBlob(base64Data, "document.pdf", "application/pdf");
 */
function base64ToBlob(base64String, filename, mimeType) {
  try {
    const data = Utilities.base64Decode(base64String);
    return Utilities.newBlob(
      data,
      mimeType || "application/octet-stream",
      filename
    );
  } catch (error) {
    logError("base64ToBlob", error);
    throw new Error("Impossible de décoder le fichier : " + filename);
  }
}

/**
 * Valide la taille d'un fichier base64
 * 
 * @param {string} base64String - Données encodées
 * @param {number} maxSizeMB - Taille max en MB
 * @returns {boolean} true si taille OK
 * @throws {Error} Si fichier trop volumineux
 */
function validateFileSize(base64String, maxSizeMB) {
  const sizeBytes = base64String.length * 0.75; // Approximation base64 -> octets
  const sizeMB = sizeBytes / (1024 * 1024);
  
  if (sizeMB > maxSizeMB) {
    throw new Error(
      `Fichier trop volumineux : ${sizeMB.toFixed(2)}MB (max: ${maxSizeMB}MB)`
    );
  }
  
  return true;
}

/**
 * ============================================
 * GESTION D'ERREURS STANDARDISÉE
 * ============================================
 */

/**
 * Formatte une erreur pour l'affichage utilisateur
 * 
 * @param {Error} error - Objet erreur
 * @param {string} context - Contexte de l'erreur
 * @returns {string} Message d'erreur formaté
 */
function formatErrorForUser(error, context) {
  const message = error.message || String(error);
  
  // Log technique complet
  logError(context, error);
  
  // Message utilisateur simplifié
  if (message.includes("quota")) {
    return "⚠️ Quota dépassé. " + message;
  }
  if (message.includes("permission")) {
    return "🔒 Erreur de permissions. Vérifiez vos accès.";
  }
  if (message.includes("timeout")) {
    return "⏱️ Délai d'attente dépassé. Réessayez dans quelques instants.";
  }
  
  // Message générique si pas de cas spécifique
  return "❌ Une erreur est survenue : " + message;
}

/**
 * ============================================
 * HELPERS DE VALIDATION
 * ============================================
 */

/**
 * Vérifie si une valeur est vide/null/undefined
 * 
 * @param {*} value - Valeur à tester
 * @returns {boolean} true si vide
 */
function isEmpty(value) {
  return value === null || 
         value === undefined || 
         value === "" || 
         (Array.isArray(value) && value.length === 0) ||
         (typeof value === "object" && Object.keys(value).length === 0);
}

/**
 * Vérifie qu'un objet a toutes les clés requises
 * 
 * @param {Object} obj - Objet à valider
 * @param {Array<string>} requiredKeys - Clés obligatoires
 * @returns {boolean} true si toutes les clés présentes
 * @throws {Error} Si clé manquante
 */
function validateRequiredFields(obj, requiredKeys) {
  if (!obj || typeof obj !== "object") {
    throw new Error("Objet invalide");
  }
  
  const missing = requiredKeys.filter(key => isEmpty(obj[key]));
  
  if (missing.length > 0) {
    throw new Error(`Champs requis manquants : ${missing.join(", ")}`);
  }
  
  return true;
}
