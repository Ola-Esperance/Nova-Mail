/*****************************************************
 * NOVAMAIL SAAS - CONFIG.GS
 * ====================================================
 * Gestion centralisée des versions, quotas et configuration
 * 
 * @author NovaMail Team
 * @version 2.0.0
 * @lastModified 2025-11-03
 *****************************************************/

/**
 * ============================================
 * CONSTANTES GLOBALES
 * ============================================
 */

// Email et identité de l'expéditeur par défaut
const DEFAULT_SENDER_EMAIL = "foreverjoyfulcreations@gmail.com";
const DEFAULT_SENDER_NAME = "NovaMail";

// Limites techniques Gmail
const GMAIL_BATCH_SIZE = 40; // Envois par batch pour éviter les timeouts
const GMAIL_BATCH_DELAY_MS = 1500; // Délai entre batches en ms
const MAX_ATTACHMENT_SIZE_MB = 15; // Taille max par pièce jointe

// Préfixes pour le stockage des données
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
 * CONFIGURATION DES VERSIONS PRODUIT
 * ============================================
 * Définit les limites et fonctionnalités par version
 */

/**
 * Récupère la configuration de la version active pour l'utilisateur courant
 * 
 * @returns {Object} Configuration de version avec limites et features
 * 
 * @example
 * const config = getVersionConfig();
 * if (config.allowAttachments) { ... }
 */
function getVersionConfig() {
  // Récupération de la version utilisateur (à terme depuis PropertiesService)
  const userVersion = getUserVersion();
  
  const versionConfigs = {
    FREE: {
      name: "Free",
      displayName: "Version Gratuite",
      maxRecipients: 10,           // Max destinataires par campagne
      monthlyQuota: 50,             // Max envois par mois
      annualQuota: 500,             // Max envois par an
      allowAttachments: false,      // Pièces jointes désactivées
      allowImportSheets: false,     // Import Google Sheets désactivé
      allowTemplateSave: false,     // Sauvegarde modèles désactivée
      multiSender: false,           // Un seul expéditeur
      scheduleSend: false,          // Pas de planification
      customBranding: false,        // Branding NovaMail obligatoire
      maxTemplates: 0,              // Pas de templates
      analyticsEnabled: false,      // Stats désactivées
      prioritySupport: false        // Support standard uniquement
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
      scheduleSend: "limited",      // Planification limitée (48h max)
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
      multiSender: true,            // Multiples expéditeurs
      scheduleSend: true,           // Planification illimitée
      customBranding: true,         // Branding personnalisable
      maxTemplates: 20,
      analyticsEnabled: true,
      prioritySupport: true,
      webhooksEnabled: true         // Webhooks pour intégrations
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
      scheduleSend: "recurring",    // Campagnes récurrentes
      customBranding: true,
      maxTemplates: -1,             // Illimité
      analyticsEnabled: true,
      prioritySupport: true,
      webhooksEnabled: true,
      apiAccess: true,              // Accès API REST
      dedicatedSupport: true,       // Support dédié
      slaGuarantee: "99.9%"         // Garantie SLA
    }
  };
  
  // Retourne la config correspondante ou FREE par défaut
  return versionConfigs[userVersion] || versionConfigs.FREE;
}

/**
 * ============================================
 * GESTION DE LA VERSION UTILISATEUR
 * ============================================
 */

/**
 * Récupère la version produit assignée à l'utilisateur courant
 * 
 * @returns {string} Code version (FREE|STARTER|PRO|BUSINESS)
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
  
  // Par défaut : version FREE
  return "FREE";
}

/**
 * Définit la version produit pour l'utilisateur courant
 * ⚠️ Fonction sécurisée : à n'appeler que lors de l'activation client
 * 
 * @param {string} version - Code version à assigner
 * @returns {boolean} Succès de l'opération
 * 
 * @example
 * setUserVersion("PRO"); // Active la version Pro pour l'utilisateur
 */
function setUserVersion(version) {
  // Validation stricte de la version
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
 * GESTION DES QUOTAS UTILISATEUR
 * ============================================
 */

/**
 * Structure des données de quota utilisateur
 * @typedef {Object} UserQuota
 * @property {number} monthly - Nombre d'envois ce mois
 * @property {number} annual - Nombre d'envois cette année
 * @property {string} lastReset - ISO timestamp du dernier reset
 * @property {string} monthYear - Référence mois/année (YYYY-MM)
 */

/**
 * Récupère les quotas de consommation de l'utilisateur
 * Initialise automatiquement si première utilisation
 * 
 * @returns {UserQuota} Objet quota utilisateur
 */
function getUserQuota() {
  try {
    const raw = PropertiesService.getUserProperties()
      .getProperty(STORAGE_PREFIX.USER_PROPS + ":" + USER_STORAGE_KEYS.USER_QUOTA);
    
    if (!raw) {
      // Initialisation quota vierge
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
    
    // Vérification reset mensuel automatique
    if (quota.monthYear !== getMonthYearKey()) {
      quota.monthly = 0;
      quota.monthYear = getMonthYearKey();
      quota.lastReset = new Date().toISOString();
      saveUserQuota(quota);
    }
    
    return quota;
    
  } catch (error) {
    logError("getUserQuota", error);
    // Retour quota vierge en cas d'erreur
    return {
      monthly: 0,
      annual: 0,
      lastReset: new Date().toISOString(),
      monthYear: getMonthYearKey()
    };
  }
}

/**
 * Sauvegarde les quotas utilisateur
 * 
 * @param {UserQuota} quota - Objet quota à sauvegarder
 * @returns {boolean} Succès de l'opération
 */
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

/**
 * Incrémente les compteurs de quota après un envoi réussi
 * 
 * @param {number} count - Nombre de destinataires envoyés
 * @returns {boolean} Succès de l'opération
 */
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

/**
 * Vérifie si l'utilisateur peut envoyer à N destinataires
 * Ne modifie PAS les quotas (juste vérification)
 * 
 * @param {number} recipientCount - Nombre de destinataires à envoyer
 * @throws {Error} Si quota dépassé
 * @returns {boolean} true si quota OK
 */
function checkQuotaAvailable(recipientCount) {
  const config = getVersionConfig();
  const quota = getUserQuota();
  
  // Vérification limite par campagne
  if (recipientCount > config.maxRecipients) {
    throw new Error(
      `🚫 Limite par campagne dépassée. Maximum : ${config.maxRecipients} destinataires ` +
      `(votre version : ${config.displayName})`
    );
  }
  
  // Vérification quota mensuel
  if (quota.monthly + recipientCount > config.monthlyQuota) {
    throw new Error(
      `🚫 Quota mensuel dépassé. Maximum : ${config.monthlyQuota} envois/mois. ` +
      `Utilisé : ${quota.monthly}. Upgrade pour plus de capacité.`
    );
  }
  
  // Vérification quota annuel
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
 * HELPERS INTERNES
 * ============================================
 */

/**
 * Génère une clé mois-année pour tracking des resets
 * @returns {string} Format YYYY-MM
 */
function getMonthYearKey() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  return `${year}-${month}`;
}

/**
 * ============================================
 * LOGGING CENTRALISÉ
 * ============================================
 */

/**
 * Log d'information
 * @param {string} message - Message à logger
 */
function logInfo(message) {
  Logger.log(`[INFO] ${new Date().toISOString()} - ${message}`);
}

/**
 * Log d'erreur avec contexte
 * @param {string} context - Nom de la fonction
 * @param {Error} error - Objet erreur
 */
function logError(context, error) {
  Logger.log(`[ERROR] ${new Date().toISOString()} - ${context}: ${error.message}`);
  console.error(error); // Stack trace complète en console
}

/**
 * Log d'avertissement
 * @param {string} message - Message à logger
 */
function logWarning(message) {
  Logger.log(`[WARNING] ${new Date().toISOString()} - ${message}`);
}
