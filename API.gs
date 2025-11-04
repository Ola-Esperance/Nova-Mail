/*****************************************************
 * NOVAMAIL SAAS - API.GS
 * ====================================================
 * Points d'entrée exposés au frontend (index.html)
 * Interface standardisée entre backend et UI
 * 
 * @author NovaMail Team
 * @version 2.0.0
 * @lastModified 2025-11-03
 * 
 * ⚠️ IMPORTANT : Ne jamais renommer ces fonctions sans mettre
 * à jour les appels correspondants dans index.html
 *****************************************************/

/**
 * ============================================
 * GESTION DES ERREURS API
 * ============================================
 */

/**
 * Wrapper standardisé pour toutes les fonctions API
 * Gère les erreurs et retourne un format cohérent
 * 
 * @param {Function} fn - Fonction à exécuter
 * @param {string} context - Nom de la fonction (pour logs)
 * @returns {*} Résultat de la fonction ou objet d'erreur
 */
function apiWrapper(fn, context) {
  try {
    return fn();
  } catch (error) {
    logError(context, error);
    return {
      success: false,
      error: error.message || String(error)
    };
  }
}

/**
 * ============================================
 * VERSION & CONFIGURATION
 * ============================================
 */

/**
 * 📋 Récupère la configuration de version de l'utilisateur courant
 * Exposé au frontend pour adapter l'UI selon les permissions
 * 
 * @returns {Object} Configuration version
 * 
 * @example (côté frontend)
 * google.script.run.withSuccessHandler(config => {
 *   if (config.allowAttachments) { ... }
 * }).getAppVersion();
 */
function getAppVersion() {
  return apiWrapper(() => {
    const config = getVersionConfig();
    const quota = getUserQuota();
    
    return {
      ...config,
      quota: {
        monthly: quota.monthly,
        annual: quota.annual,
        monthlyLimit: config.monthlyQuota,
        annualLimit: config.annualQuota,
        monthlyRemaining: config.monthlyQuota - quota.monthly,
        annualRemaining: config.annualQuota - quota.annual
      }
    };
  }, "getAppVersion");
}

/**
 * ============================================
 * IMPORT DE DESTINATAIRES
 * ============================================
 */

/**
 * 📊 Liste les Google Sheets accessibles
 * 
 * @returns {Array<Object>} Liste des sheets
 */
function listGoogleSheets() {
  return apiWrapper(() => {
    return listGoogleSheets(); // Appel fonction dans Campaigns.gs
  }, "listGoogleSheets");
}

/**
 * 📄 Importe depuis un Google Sheet
 * 
 * @param {string} fileId - ID du fichier
 * @param {string} sheetName - Nom de la feuille (optionnel)
 * @returns {Array<Object>} Destinataires
 */
function importFromSheet(fileId, sheetName) {
  return apiWrapper(() => {
    return importFromSheet(fileId, sheetName);
  }, "importFromSheet");
}

/**
 * 📄 Importe depuis un fichier CSV
 * 
 * @param {string} csvContent - Contenu CSV (texte ou base64)
 * @returns {Array<Object>} Destinataires
 */
function importCsv(csvContent) {
  return apiWrapper(() => {
    return importCsv(csvContent);
  }, "importCsv");
}

/**
 * ⤴️ Récupère la dernière liste importée
 * 
 * @returns {Array<Object>} Dernière liste
 */
function getLastImportedList() {
  return apiWrapper(() => {
    return getLastImportedList();
  }, "getLastImportedList");
}

/**
 * 🧹 Efface la dernière liste importée
 * 
 * @returns {boolean} Succès
 */
function clearLastImportedList() {
  return apiWrapper(() => {
    return clearLastImportedList();
  }, "clearLastImportedList");
}

/**
 * ============================================
 * GESTION DES TEMPLATES
 * ============================================
 */

/**
 * 💾 Sauvegarde un template
 * 
 * @param {Object} template - Objet {name, subject, htmlBody}
 * @returns {Object} Template sauvegardé
 */
function saveTemplate(template) {
  return apiWrapper(() => {
    if (!template || !template.name) {
      throw new Error("Nom du template requis");
    }
    
    return saveTemplate(template.name, template.subject, template.htmlBody);
  }, "saveTemplate");
}

/**
 * 📄 Charge un template par nom
 * 
 * @param {string} name - Nom du template
 * @returns {Object|null} Template
 */
function loadTemplate(name) {
  return apiWrapper(() => {
    return loadTemplate(name);
  }, "loadTemplate");
}

/**
 * 📚 Liste tous les templates
 * 
 * @returns {Array<Object>} Liste des templates
 */
function listTemplates() {
  return apiWrapper(() => {
    return listTemplates();
  }, "listTemplates");
}

/**
 * 🗑️ Supprime un template
 * 
 * @param {string} name - Nom du template
 * @returns {Object} Résultat {success}
 */
function deleteTemplate(name) {
  return apiWrapper(() => {
    deleteTemplate(name);
    return { success: true, name: name };
  }, "deleteTemplate");
}

/**
 * 🧹 Supprime tous les templates
 * 
 * @returns {Object} Résultat {success, count}
 */
function deleteAllTemplates() {
  return apiWrapper(() => {
    const count = deleteAllTemplates();
    return { success: true, count: count };
  }, "deleteAllTemplates");
}

/**
 * ============================================
 * PRÉVISUALISATION
 * ============================================
 */

/**
 * 🔍 Génère une prévisualisation d'email
 * 
 * @param {string} subject - Sujet
 * @param {string} htmlBody - Contenu HTML
 * @param {Object} recipient - Destinataire exemple
 * @returns {Object} Prévisualisation
 */
function previewEmail(subject, htmlBody, recipient) {
  return apiWrapper(() => {
    return previewEmail(subject, htmlBody, recipient);
  }, "previewEmail");
}

/**
 * ============================================
 * ENVOI DE CAMPAGNES
 * ============================================
 */

/**
 * 📤 Envoie une campagne directe
 * 
 * @param {Array<Object>} recipients - Destinataires
 * @param {string} subject - Objet
 * @param {string} htmlBody - Contenu HTML
 * @param {Array<Object>} attachments - Pièces jointes
 * @returns {Object} Résultat {success, sent, failed}
 */
function sendCampaign(recipients, subject, htmlBody, attachments) {
  return apiWrapper(() => {
    return sendCampaign(recipients, subject, htmlBody, attachments);
  }, "sendCampaign");
}

/**
 * ✉️ Envoie un email de test à l'utilisateur courant
 * 
 * @param {string} subject - Objet
 * @param {string} htmlBody - Contenu HTML
 * @param {Array<Object>} attachments - Pièces jointes
 * @returns {string} Message de confirmation
 */
function sendTestToMe(subject, htmlBody, attachments) {
  return apiWrapper(() => {
    return sendTestToMe(subject, htmlBody, attachments);
  }, "sendTestToMe");
}

/**
 * ============================================
 * PLANIFICATION
 * ============================================
 */

/**
 * 📅 Planifie une campagne
 * 
 * @param {Object} campaign - Objet campagne avec sendAt
 * @returns {Object} Résultat {success, id, scheduledAt}
 */
function scheduleCampaign(campaign) {
  return apiWrapper(() => {
    return scheduleCampaign(campaign);
  }, "scheduleCampaign");
}

/**
 * 📋 Liste les campagnes planifiées
 * 
 * @returns {Array<Object>} Liste des campagnes
 */
function getScheduledCampaigns() {
  return apiWrapper(() => {
    return getScheduledCampaigns();
  }, "getScheduledCampaigns");
}

/**
 * 🔄 Modifie la date d'une campagne planifiée
 * 
 * @param {string} campaignId - ID campagne
 * @param {string} newSendAtIso - Nouvelle date ISO
 * @returns {Object} Résultat {success, scheduledAt}
 */
function updateScheduledDate(campaignId, newSendAtIso) {
  return apiWrapper(() => {
    return updateScheduledDate(campaignId, newSendAtIso);
  }, "updateScheduledDate");
}

/**
 * 🗑️ Supprime une campagne planifiée
 * 
 * @param {string} campaignId - ID campagne
 * @returns {Object} Résultat {success}
 */
function deleteScheduledCampaign(campaignId) {
  return apiWrapper(() => {
    return deleteScheduledCampaign(campaignId);
  }, "deleteScheduledCampaign");
}

/**
 * ============================================
 * HISTORIQUE
 * ============================================
 */

/**
 * 📤 Récupère l'historique des campagnes
 * 
 * @param {number} limit - Nombre max de lignes (optionnel)
 * @returns {Array<Object>} Historique
 */
function getHistoryData(limit) {
  return apiWrapper(() => {
    return getHistoryData(limit);
  }, "getHistoryData");
}

/**
 * 📊 Récupère les statistiques d'historique
 * 
 * @returns {Object} Statistiques
 */
function getHistoryStats() {
  return apiWrapper(() => {
    return getHistoryStats();
  }, "getHistoryStats");
}

/**
 * 🧹 Efface tout l'historique
 * 
 * @returns {Object} Résultat {success}
 */
function clearHistorySheet() {
  return apiWrapper(() => {
    clearHistorySheet();
    return { success: true };
  }, "clearHistorySheet");
}

/**
 * 🧹 Nettoie les entrées de test
 * 
 * @returns {Object} Résultat {success, count}
 */
function clearTestHistory() {
  return apiWrapper(() => {
    const count = clearTestHistory();
    return { success: true, count: count };
  }, "clearTestHistory");
}

/**
 * 📥 Exporte l'historique en CSV
 * 
 * @returns {string} Contenu CSV
 */
function exportHistoryToCSV() {
  return apiWrapper(() => {
    return exportHistoryToCSV();
  }, "exportHistoryToCSV");
}

/**
 * ============================================
 * GESTION UTILISATEUR
 * ============================================
 */

/**
 * 👤 Récupère les informations de l'utilisateur connecté
 * 
 * @returns {Object} Infos utilisateur ou null
 */
function getCurrentUserInfo() {
  return apiWrapper(() => {
    const email = Session.getActiveUser().getEmail();
    
    if (!email || !isValidEmail(email)) {
      return null;
    }
    
    const client = findClientByEmail(email);
    
    if (client) {
      return {
        userId: client.userId,
        fullName: client.fullName,
        email: client.loginEmail,
        company: client.companyName,
        version: client.version,
        activatedAt: client.activatedAt,
        status: client.status
      };
    }
    
    // Utilisateur non enregistré comme client
    return {
      email: email,
      version: getUserVersion(),
      registered: false
    };
  }, "getCurrentUserInfo");
}

/**
 * ============================================
 * FONCTIONS UTILITAIRES
 * ============================================
 */

/**
 * 🧪 Teste la connexion et les permissions
 * 
 * @returns {Object} Résultat des tests
 */
function testConnection() {
  return apiWrapper(() => {
    const tests = {
      success: true,
      checks: []
    };
    
    // Test 1 : Lecture email utilisateur
    try {
      const email = Session.getActiveUser().getEmail();
      tests.checks.push({
        name: "Email utilisateur",
        passed: !!email,
        value: email || "Non détecté"
      });
    } catch (e) {
      tests.checks.push({
        name: "Email utilisateur",
        passed: false,
        error: e.message
      });
      tests.success = false;
    }
    
    // Test 2 : Permission Gmail
    try {
      GmailApp.getAliases();
      tests.checks.push({
        name: "Permission Gmail",
        passed: true,
        value: "OK"
      });
    } catch (e) {
      tests.checks.push({
        name: "Permission Gmail",
        passed: false,
        error: e.message
      });
      tests.success = false;
    }
    
    // Test 3 : Accès Script Properties
    try {
      PropertiesService.getScriptProperties().getProperty("TEST");
      tests.checks.push({
        name: "Script Properties",
        passed: true,
        value: "OK"
      });
    } catch (e) {
      tests.checks.push({
        name: "Script Properties",
        passed: false,
        error: e.message
      });
      tests.success = false;
    }
    
    // Test 4 : Configuration version
    try {
      const config = getVersionConfig();
      tests.checks.push({
        name: "Configuration version",
        passed: true,
        value: config.name
      });
    } catch (e) {
      tests.checks.push({
        name: "Configuration version",
        passed: false,
        error: e.message
      });
      tests.success = false;
    }
    
    return tests;
  }, "testConnection");
}

/**
 * 🔧 Récupère des informations de debug
 * 
 * @returns {Object} Infos de debug
 */
function getDebugInfo() {
  return apiWrapper(() => {
    return {
      scriptId: ScriptApp.getScriptId(),
      timezone: Session.getScriptTimeZone(),
      userEmail: Session.getActiveUser().getEmail(),
      version: getUserVersion(),
      deploymentId: getDeploymentId(),
      historyFileId: PropertiesService.getScriptProperties()
        .getProperty(HISTORY_CONFIG.PROPERTY_KEY),
      triggersCount: ScriptApp.getProjectTriggers().length
    };
  }, "getDebugInfo");
}

/**
 * ============================================
 * POINT D'ENTRÉE WEB (doGet)
 * ============================================
 */

/**
 * 🌐 Point d'entrée HTTP pour l'application web
 * Déjà défini dans UserManagement.gs, on le laisse tel quel
 * 
 * Cette fonction gère :
 * - Accès sans userId : interface générique
 * - Accès avec userId : interface personnalisée client
 * 
 * @param {Object} e - Event object
 * @returns {HtmlOutput} Page HTML
 */
// function doGet(e) est déjà défini dans UserManagement.gs

/**
 * ============================================
 * HELPERS POUR LE FRONTEND
 * ============================================
 */

/**
 * 📋 Récupère toutes les données nécessaires au chargement de l'interface
 * Optimisation : un seul appel depuis le frontend au lieu de plusieurs
 * 
 * @returns {Object} Données complètes
 * 
 * @example (côté frontend)
 * google.script.run.withSuccessHandler(data => {
 *   // data.config, data.quota, data.user, data.templates, etc.
 * }).loadAppData();
 */
function loadAppData() {
  return apiWrapper(() => {
    return {
      config: getVersionConfig(),
      quota: getUserQuota(),
      user: getCurrentUserInfo(),
      templates: listTemplates(),
      scheduledCampaigns: getScheduledCampaigns(),
      historyStats: getHistoryStats(),
      lastImportedList: getLastImportedList()
    };
  }, "loadAppData");
}

/**
 * ============================================
 * FONCTIONS DE MAINTENANCE
 * ============================================
 */

/**
 * 🔄 Réinitialise l'application (⚠️ admin seulement)
 * 
 * @param {string} confirmationCode - Code de confirmation
 * @returns {Object} Résultat
 */
function resetApplication(confirmationCode) {
  return apiWrapper(() => {
    // Code de sécurité
    if (confirmationCode !== "RESET_NOVAMAIL_2025") {
      throw new Error("Code de confirmation invalide");
    }
    
    // Suppression triggers
    removeMasterTrigger();
    removeAllTimeDrivenTriggers();
    
    // Nettoyage historique
    clearHistorySheet();
    
    // Note : les données clients et quotas sont conservés
    
    logWarning("⚠️ Application réinitialisée (données clients conservées)");
    
    return {
      success: true,
      message: "Application réinitialisée. Données clients conservées."
    };
  }, "resetApplication");
}

/**
 * 📊 Récupère un rapport complet du système
 * 
 * @returns {Object} Rapport système
 */
function getSystemReport() {
  return apiWrapper(() => {
    const triggers = ScriptApp.getProjectTriggers();
    
    return {
      version: "2.0.0",
      environment: {
        scriptId: ScriptApp.getScriptId(),
        timezone: Session.getScriptTimeZone(),
        quotaRemaining: getQuotaRemaining()
      },
      triggers: {
        total: triggers.length,
        scheduled: triggers.filter(t => 
          t.getHandlerFunction() === SCHEDULING_CONFIG.MASTER_TRIGGER_FUNCTION
        ).length,
        sheets: triggers.filter(t => 
          t.getHandlerFunction() === TRIGGER_CONFIG.HANDLER_FUNCTION
        ).length
      },
      clients: {
        total: Object.keys(loadClientIndex()).length
      },
      campaigns: {
        scheduled: getScheduledCampaigns().length
      },
      history: getHistoryStats(),
      status: "operational"
    };
  }, "getSystemReport");
}

/**
 * Calcule le quota restant (approximatif)
 * 
 * @returns {number} Estimation quota Google Apps Script
 */
function getQuotaRemaining() {
  try {
    // Approximation basée sur le temps d'exécution
    const start = new Date().getTime();
    Utilities.sleep(10);
    const elapsed = new Date().getTime() - start;
    
    // Quota Apps Script : ~6min d'exécution / jour
    // Retour approximatif en secondes
    return Math.max(0, 360 - elapsed / 1000);
  } catch (e) {
    return -1;
  }
}
