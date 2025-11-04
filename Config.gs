/*****************************************************
 * NOVAMAIL SAAS - CONFIG.GS (VERSION CORRIGÉE)
 * ====================================================
 * ✅ FIX : DEFAULT_SENDER_EMAIL auto-détecté
 * ✅ FIX : SHEET_ID configurable au lieu du parcours Drive
 * 
 * @version 2.1.0 FIXED
 * @lastModified 2025-11-04
 *****************************************************/

/**
 * ============================================
 * CONSTANTES GLOBALES (AUTO-DÉTECTION)
 * ============================================
 */

/**
 * 📧 Email expéditeur par défaut
 * AUTO-DÉTECTÉ depuis le compte Apps Script actif
 * Peut être surchargé via setDefaultSenderEmail()
 */
function getDefaultSenderEmail() {
  try {
    // Tentative 1 : Récupérer depuis Script Properties (si configuré manuellement)
    const stored = PropertiesService.getScriptProperties()
      .getProperty("DEFAULT_SENDER_EMAIL");
    
    if (stored && isValidEmail(stored)) {
      return stored;
    }
    
    // Tentative 2 : Email du compte actif (développeur ou compte de service)
    const activeEmail = Session.getActiveUser().getEmail();
    
    if (activeEmail && isValidEmail(activeEmail)) {
      // Mémorisation pour performances
      PropertiesService.getScriptProperties()
        .setProperty("DEFAULT_SENDER_EMAIL", activeEmail);
      
      logInfo(`📧 DEFAULT_SENDER_EMAIL auto-détecté : ${activeEmail}`);
      return activeEmail;
    }
    
    // Tentative 3 : Effective user (compte de service)
    const effectiveEmail = Session.getEffectiveUser().getEmail();
    
    if (effectiveEmail && isValidEmail(effectiveEmail)) {
      PropertiesService.getScriptProperties()
        .setProperty("DEFAULT_SENDER_EMAIL", effectiveEmail);
      
      logInfo(`📧 DEFAULT_SENDER_EMAIL (effective) : ${effectiveEmail}`);
      return effectiveEmail;
    }
    
    // ⚠️ Fallback ultime : email depuis DEV_CONFIG
    const devEmail = DEV_CONFIG.email;
    
    if (devEmail && isValidEmail(devEmail)) {
      logWarning(`⚠️ Utilisation email développeur comme fallback : ${devEmail}`);
      return devEmail;
    }
    
    // ❌ Cas extrême : aucun email trouvé
    throw new Error(
      "Impossible de déterminer DEFAULT_SENDER_EMAIL. " +
      "Configurez-le manuellement via setDefaultSenderEmail()"
    );
    
  } catch (error) {
    logError("getDefaultSenderEmail", error);
    
    // Retour email développeur comme dernier recours
    return DEV_CONFIG.email || "noreply@novamail.app";
  }
}

/**
 * Définit manuellement l'email expéditeur par défaut
 * 
 * @param {string} email - Email à définir
 * @returns {boolean} Succès
 * 
 * @example
 * setDefaultSenderEmail("support@monentreprise.com");
 */
function setDefaultSenderEmail(email) {
  if (!email || !isValidEmail(email)) {
    logError("setDefaultSenderEmail", new Error("Email invalide : " + email));
    return false;
  }
  
  try {
    PropertiesService.getScriptProperties()
      .setProperty("DEFAULT_SENDER_EMAIL", email);
    
    logInfo(`✅ DEFAULT_SENDER_EMAIL configuré : ${email}`);
    return true;
    
  } catch (error) {
    logError("setDefaultSenderEmail", error);
    return false;
  }
}

/**
 * ⚠️ RÉTRO-COMPATIBILITÉ : Constante pour usages legacy
 * Utilise la fonction pour éviter les erreurs de définition
 */
const DEFAULT_SENDER_EMAIL = getDefaultSenderEmail();

/**
 * Nom de l'expéditeur par défaut
 */
const DEFAULT_SENDER_NAME = "NovaMail";

/**
 * ============================================
 * GOOGLE SHEET ID (FIX PARCOURS DRIVE)
 * ============================================
 */

/**
 * 📊 Récupère l'ID du Google Sheet à utiliser
 * PLUS DE PARCOURS DRIVE — Configuration directe via ID
 * 
 * @returns {string|null} ID du spreadsheet ou null
 */
function getSourceSheetId() {
  try {
    // Méthode 1 : ID stocké dans Script Properties
    const storedId = PropertiesService.getScriptProperties()
      .getProperty("SOURCE_SHEET_ID");
    
    if (storedId) {
      logInfo(`📊 Sheet ID récupéré : ${storedId}`);
      return storedId;
    }
    
    // Méthode 2 : Script lié à un spreadsheet (container-bound)
    try {
      const boundSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      if (boundSpreadsheet && boundSpreadsheet.getId()) {
        const id = boundSpreadsheet.getId();
        
        // Enregistrer pour usage futur
        PropertiesService.getScriptProperties()
          .setProperty("SOURCE_SHEET_ID", id);
        
        logInfo(`📊 Sheet ID auto-détecté (container-bound) : ${id}`);
        return id;
      }
    } catch (e) {
      // Pas un script container-bound
    }
    
    // ⚠️ Pas de sheet configuré
    logWarning("⚠️ Aucun Sheet ID configuré. Utilisez setSourceSheetId()");
    return null;
    
  } catch (error) {
    logError("getSourceSheetId", error);
    return null;
  }
}

/**
 * Configure manuellement l'ID du Google Sheet source
 * 
 * @param {string} sheetId - ID du spreadsheet
 * @returns {boolean} Succès
 * 
 * @example
 * setSourceSheetId("1abc...XYZ");
 */
function setSourceSheetId(sheetId) {
  if (!sheetId || typeof sheetId !== "string") {
    logError("setSourceSheetId", new Error("Sheet ID invalide"));
    return false;
  }
  
  try {
    // Validation : tenter d'ouvrir le sheet
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    
    if (!spreadsheet) {
      throw new Error("Sheet inaccessible ou ID invalide");
    }
    
    // Enregistrement
    PropertiesService.getScriptProperties()
      .setProperty("SOURCE_SHEET_ID", sheetId);
    
    logInfo(`✅ Sheet ID configuré : ${sheetId} (${spreadsheet.getName()})`);
    return true;
    
  } catch (error) {
    logError("setSourceSheetId", error);
    return false;
  }
}

/**
 * Ouvre directement le Google Sheet configuré
 * 
 * @returns {Spreadsheet|null} Spreadsheet ou null
 */
function openSourceSheet() {
  const sheetId = getSourceSheetId();
  
  if (!sheetId) {
    logError("openSourceSheet", new Error("Sheet ID non configuré"));
    return null;
  }
  
  try {
    return SpreadsheetApp.openById(sheetId);
  } catch (error) {
    logError("openSourceSheet", error);
    return null;
  }
}

/**
 * ============================================
 * RESTE DE LA CONFIGURATION (INCHANGÉ)
 * ============================================
 */

// Limites techniques Gmail
const GMAIL_BATCH_SIZE = 40;
const GMAIL_BATCH_DELAY_MS = 1500;
const MAX_ATTACHMENT_SIZE_MB = 15;

// Préfixes pour le stockage
const STORAGE_PREFIX = {
  USER_PROPS: "NOVAMAIL_USER",
  SCRIPT_PROPS: "NOVAMAIL_SCRIPT",
  SCHEDULED_CAMPAIGN: "SCHEDULED_",
  TEMPLATE: "TEMPLATE:"
};

// Gestion de l'historique
const HISTORY_CONFIG = {
  SHEET_NAME: "Historique_Campagnes",
  FILE_NAME: "NovaMail_Historique",
  PROPERTY_KEY: "HISTORY_SPREADSHEET_ID"
};

// Clés de stockage utilisateur
const USER_STORAGE_KEYS = {
  LAST_LIST: "LAST_LIST",
  USER_QUOTA: "USER_QUOTA",
  USER_VERSION: "USER_VERSION"
};

/**
 * ============================================
 * CONFIGURATION DES VERSIONS (INCHANGÉ)
 * ============================================
 */

function getVersionConfig() {
  const userVersion = getUserVersion();
  
  const versionConfigs = {
    FREE: {
      name: "Free",
      displayName: "Version Gratuite",
      maxRecipients: 10,
      monthlyQuota: 50,
      annualQuota: 500,
      allowAttachments: false,
      allowImportSheets: false,
      allowTemplateSave: false,
      multiSender: false,
      scheduleSend: false,
      customBranding: false,
      maxTemplates: 0,
      analyticsEnabled: false,
      prioritySupport: false
    },
    
    STARTER: {
      name: "Starter",
      displayName: "Version Starter",
      maxRecipients: 200,
      monthlyQuota: 2000,
      annualQuota: 20000,
      allowAttachments: true,
      allowImportSheets: true,
      allowTemplateSave: true,
      multiSender: false,
      scheduleSend: "limited",
      customBranding: false,
      maxTemplates: 5,
      analyticsEnabled: true,
      prioritySupport: false
    },
    
    PRO: {
      name: "Pro",
      displayName: "Version Professionnelle",
      maxRecipients: 1000,
      monthlyQuota: 10000,
      annualQuota: 120000,
      allowAttachments: true,
      allowImportSheets: true,
      allowTemplateSave: true,
      multiSender: true,
      scheduleSend: true,
      customBranding: true,
      maxTemplates: 20,
      analyticsEnabled: true,
      prioritySupport: true,
      webhooksEnabled: true
    },
    
    BUSINESS: {
      name: "Business",
      displayName: "Version Entreprise",
      maxRecipients: 5000,
      monthlyQuota: 50000,
      annualQuota: 600000,
      allowAttachments: true,
      allowImportSheets: true,
      allowTemplateSave: true,
      multiSender: true,
      scheduleSend: "recurring",
      customBranding: true,
      maxTemplates: -1,
      analyticsEnabled: true,
      prioritySupport: true,
      webhooksEnabled: true,
      apiAccess: true,
      dedicatedSupport: true,
      slaGuarantee: "99.9%"
    }
  };
  
  return versionConfigs[userVersion] || versionConfigs.FREE;
}

/**
 * ============================================
 * GESTION VERSION UTILISATEUR (INCHANGÉ)
 * ============================================
 */

function getUserVersion() {
  try {
    const stored = PropertiesService.getUserProperties()
      .getProperty(STORAGE_PREFIX.USER_PROPS + ":" + USER_STORAGE_KEYS.USER_VERSION);
    
    if (stored && ["FREE", "STARTER", "PRO", "BUSINESS"].includes(stored.toUpperCase())) {
      return stored.toUpperCase();
    }
  } catch (error) {
    logError("getUserVersion", error);
  }
  
  return "FREE";
}

function setUserVersion(version) {
  const validVersions = ["FREE", "STARTER", "PRO", "BUSINESS"];
  const normalizedVersion = String(version).toUpperCase().trim();
  
  if (!validVersions.includes(normalizedVersion)) {
    logError("setUserVersion", new Error("Version invalide: " + version));
    return false;
  }
  
  try {
    PropertiesService.getUserProperties().setProperty(
      STORAGE_PREFIX.USER_PROPS + ":" + USER_STORAGE_KEYS.USER_VERSION,
      normalizedVersion
    );
    
    logInfo("Version utilisateur mise à jour: " + normalizedVersion);
    return true;
  } catch (error) {
    logError("setUserVersion", error);
    return false;
  }
}

/**
 * ============================================
 * GESTION QUOTAS (INCHANGÉ)
 * ============================================
 */

function getUserQuota() {
  try {
    const raw = PropertiesService.getUserProperties()
      .getProperty(STORAGE_PREFIX.USER_PROPS + ":" + USER_STORAGE_KEYS.USER_QUOTA);
    
    if (!raw) {
      const initQuota = {
        monthly: 0,
        annual: 0,
        lastReset: new Date().toISOString(),
        monthYear: getMonthYearKey()
      };
      saveUserQuota(initQuota);
      return initQuota;
    }
    
    const quota = JSON.parse(raw);
    
    if (quota.monthYear !== getMonthYearKey()) {
      quota.monthly = 0;
      quota.monthYear = getMonthYearKey();
      quota.lastReset = new Date().toISOString();
      saveUserQuota(quota);
    }
    
    return quota;
    
  } catch (error) {
    logError("getUserQuota", error);
    return {
      monthly: 0,
      annual: 0,
      lastReset: new Date().toISOString(),
      monthYear: getMonthYearKey()
    };
  }
}

function saveUserQuota(quota) {
  try {
    PropertiesService.getUserProperties().setProperty(
      STORAGE_PREFIX.USER_PROPS + ":" + USER_STORAGE_KEYS.USER_QUOTA,
      JSON.stringify(quota)
    );
    return true;
  } catch (error) {
    logError("saveUserQuota", error);
    return false;
  }
}

function incrementQuota(count) {
  if (!count || count <= 0) return false;
  
  try {
    const quota = getUserQuota();
    quota.monthly += count;
    quota.annual += count;
    return saveUserQuota(quota);
  } catch (error) {
    logError("incrementQuota", error);
    return false;
  }
}

function checkQuotaAvailable(recipientCount) {
  const config = getVersionConfig();
  const quota = getUserQuota();
  
  if (recipientCount > config.maxRecipients) {
    throw new Error(
      `🚫 Limite par campagne dépassée. Maximum : ${config.maxRecipients} destinataires ` +
      `(votre version : ${config.displayName})`
    );
  }
  
  if (quota.monthly + recipientCount > config.monthlyQuota) {
    throw new Error(
      `🚫 Quota mensuel dépassé. Maximum : ${config.monthlyQuota} envois/mois. ` +
      `Utilisé : ${quota.monthly}. Upgrade pour plus de capacité.`
    );
  }
  
  if (quota.annual + recipientCount > config.annualQuota) {
    throw new Error(
      `🚫 Quota annuel dépassé. Maximum : ${config.annualQuota} envois/an. ` +
      `Utilisé : ${quota.annual}. Contactez-nous pour étendre votre forfait.`
    );
  }
  
  return true;
}

/**
 * ============================================
 * HELPERS (INCHANGÉ)
 * ============================================
 */

function getMonthYearKey() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  return `${year}-${month}`;
}

function logInfo(message) {
  Logger.log(`[INFO] ${new Date().toISOString()} - ${message}`);
}

function logError(context, error) {
  Logger.log(`[ERROR] ${new Date().toISOString()} - ${context}: ${error.message}`);
  console.error(error);
}

function logWarning(message) {
  Logger.log(`[WARNING] ${new Date().toISOString()} - ${message}`);
}
