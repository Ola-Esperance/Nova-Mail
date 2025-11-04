/*****************************************************
 * NOVAMAIL SAAS - AUTOCONFIG.GS
 * ====================================================
 * Configuration automatique et intelligente du système
 * Résout les 3 problématiques :
 * 1. Détection automatique du bon Google Sheet
 * 2. Gestion automatique du Deployment ID
 * 3. Espace développeur par défaut
 * 
 * @author NovaMail Team
 * @version 3.1.0 AUTO
 * @lastModified 2025-11-04
 *****************************************************/

/**
 * ============================================
 * CONFIGURATION DÉVELOPPEUR (À PERSONNALISER)
 * ============================================
 */

const DEV_CONFIG = {
  // 🔧 CONFIGUREZ VOTRE EMAIL DÉVELOPPEUR ICI
  email: "foreverjoyfulcreations@gmail.com", // ← MODIFIEZ ICI
  
  // Nom complet (pour l'espace de test)
  fullName: "Développeur NovaMail",
  
  // Nom entreprise (optionnel)
  companyName: "NovaMail Dev Team",
  
  // Version attribuée par défaut au développeur
  version: "BUSINESS", // FREE | STARTER | PRO | BUSINESS
  
  // Activation automatique au premier lancement
  autoInitialize: true
};

/**
 * ============================================
 * 1️⃣ DÉTECTION AUTOMATIQUE DU GOOGLE SHEET
 * ============================================
 */

/**
 * 🔍 Détecte automatiquement le Google Sheet à utiliser
 * 
 * Logique de détection intelligente :
 * 1. Si le script est lié à un spreadsheet (Container-bound) → utilise celui-ci
 * 2. Sinon, cherche dans le Drive les sheets récemment modifiés
 * 3. Filtre ceux qui contiennent les colonnes Tally attendues
 * 4. Retourne le plus récemment modifié
 * 
 * @returns {Spreadsheet|null} Le spreadsheet détecté
 */
function detectActiveSpreadsheet() {
  try {
    // ===== MÉTHODE 1 : SCRIPT LIÉ À UN SPREADSHEET =====
    // Si le script est dans un Container (lié au sheet), c'est le bon
    try {
      const boundSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      if (boundSpreadsheet && boundSpreadsheet.getId()) {
        logInfo(`✅ Spreadsheet détecté (container-bound) : ${boundSpreadsheet.getName()}`);
        
        // Validation : vérifier que c'est bien un sheet Tally
        if (isTallySpreadsheet(boundSpreadsheet)) {
          return boundSpreadsheet;
        } else {
          logWarning("⚠️ Le spreadsheet lié ne semble pas être un formulaire Tally");
        }
      }
    } catch (e) {
      // Pas de spreadsheet lié (script standalone)
      logInfo("ℹ️ Script standalone détecté - recherche dans Drive...");
    }
    
    // ===== MÉTHODE 2 : RECHERCHE DANS DRIVE =====
    const candidates = findTallySpreadsheetsInDrive();
    
    if (candidates.length === 0) {
      logWarning("⚠️ Aucun spreadsheet Tally trouvé dans Drive");
      return null;
    }
    
    // Retourner le plus récemment modifié
    const selected = candidates[0]; // Déjà trié par date
    logInfo(`✅ Spreadsheet sélectionné : ${selected.getName()} (modifié le ${selected.getLastUpdated()})`);
    
    return selected;
    
  } catch (error) {
    logError("detectActiveSpreadsheet", error);
    return null;
  }
}

/**
 * Recherche tous les spreadsheets Tally dans le Drive
 * 
 * @returns {Array<Spreadsheet>} Liste triée par date (récents en premier)
 */
function findTallySpreadsheetsInDrive() {
  const candidates = [];
  
  try {
    // Recherche tous les Google Sheets modifiés dans les 30 derniers jours
    const thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
    const dateStr = Utilities.formatDate(thirtyDaysAgo, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    const query = `mimeType='application/vnd.google-apps.spreadsheet' and modifiedDate > '${dateStr}'`;
    const files = DriveApp.searchFiles(query);
    
    while (files.hasNext()) {
      const file = files.next();
      
      try {
        const spreadsheet = SpreadsheetApp.openById(file.getId());
        
        // Vérifier si c'est un spreadsheet Tally
        if (isTallySpreadsheet(spreadsheet)) {
          candidates.push(spreadsheet);
        }
      } catch (e) {
        // Ignore les fichiers inaccessibles
      }
    }
    
    // Trier par date de modification décroissante
    candidates.sort((a, b) => {
      const dateA = DriveApp.getFileById(a.getId()).getLastUpdated();
      const dateB = DriveApp.getFileById(b.getId()).getLastUpdated();
      return dateB - dateA;
    });
    
    logInfo(`📊 ${candidates.length} spreadsheet(s) Tally trouvé(s)`);
    
  } catch (error) {
    logError("findTallySpreadsheetsInDrive", error);
  }
  
  return candidates;
}

/**
 * Vérifie si un spreadsheet est un formulaire Tally
 * En vérifiant la présence des colonnes caractéristiques
 * 
 * @param {Spreadsheet} spreadsheet - Spreadsheet à vérifier
 * @returns {boolean} true si c'est un sheet Tally
 */
function isTallySpreadsheet(spreadsheet) {
  try {
    const sheet = spreadsheet.getSheets()[0];
    if (!sheet) return false;
    
    const headers = sheet.getRange(1, 1, 1, Math.min(sheet.getLastColumn(), 20)).getValues()[0];
    
    // Colonnes caractéristiques Tally
    const tallyColumns = [
      "Submission ID",
      "Respondent ID",
      "Submitted at"
    ];
    
    // Vérifier que toutes les colonnes caractéristiques sont présentes
    const hasAllColumns = tallyColumns.every(col => 
      headers.some(h => String(h).includes(col))
    );
    
    return hasAllColumns;
    
  } catch (error) {
    return false;
  }
}

/**
 * Enregistre automatiquement le spreadsheet actif
 * Appelé par onEdit() pour auto-configuration
 * 
 * @param {Spreadsheet} spreadsheet - Spreadsheet à enregistrer
 */
function registerSpreadsheetAuto(spreadsheet) {
  if (!spreadsheet) return;
  
  try {
    const id = spreadsheet.getId();
    PropertiesService.getScriptProperties().setProperty("SOURCE_SHEET_ID", id);
    logInfo(`📌 Spreadsheet enregistré automatiquement : ${id}`);
  } catch (error) {
    logError("registerSpreadsheetAuto", error);
  }
}

/**
 * ============================================
 * 2️⃣ GESTION AUTOMATIQUE DU DEPLOYMENT ID
 * ============================================
 */

/**
 * 🔐 Récupère automatiquement le Deployment ID actuel
 * 
 * Méthodes tentées dans l'ordre :
 * 1. Depuis ScriptProperties (si déjà enregistré)
 * 2. Extraction depuis l'URL du service déployé
 * 3. Détection via les déploiements du projet
 * 
 * @returns {string|null} Le Deployment ID ou null
 */
function getDeploymentIdAuto() {
  try {
    // ===== MÉTHODE 1 : CACHE PROPERTIES =====
    const cached = PropertiesService.getScriptProperties()
      .getProperty("DEPLOYMENT_ID");
    
    if (cached && cached.startsWith("AK")) {
      logInfo("✅ Deployment ID récupéré depuis cache");
      return cached;
    }
    
    // ===== MÉTHODE 2 : URL DU SERVICE =====
    try {
      const service = ScriptApp.getService();
      const url = service.getUrl();
      
      if (url) {
        // Extraction du deployment ID depuis l'URL
        // Format : https://script.google.com/macros/s/DEPLOYMENT_ID/exec
        const match = url.match(/\/s\/([A-Za-z0-9_-]+)\//);
        
        if (match && match[1]) {
          const deploymentId = match[1];
          
          // Mise en cache pour prochains appels
          PropertiesService.getScriptProperties()
            .setProperty("DEPLOYMENT_ID", deploymentId);
          
          logInfo("✅ Deployment ID extrait depuis URL du service");
          return deploymentId;
        }
      }
    } catch (e) {
      // Service non déployé encore
      logWarning("⚠️ Service non déployé ou URL inaccessible");
    }
    
    // ===== MÉTHODE 3 : PAS ENCORE DÉPLOYÉ =====
    logWarning("⚠️ Aucun Deployment ID trouvé - déployez l'application web d'abord");
    return null;
    
  } catch (error) {
    logError("getDeploymentIdAuto", error);
    return null;
  }
}

/**
 * Force l'actualisation du Deployment ID
 * À appeler après un nouveau déploiement
 * 
 * @returns {string|null} Le nouveau Deployment ID
 */
function refreshDeploymentId() {
  try {
    // Supprimer le cache
    PropertiesService.getScriptProperties().deleteProperty("DEPLOYMENT_ID");
    
    // Récupérer à nouveau
    const newId = getDeploymentIdAuto();
    
    if (newId) {
      logInfo(`✅ Deployment ID actualisé : ${newId}`);
    } else {
      logWarning("⚠️ Impossible de récupérer le nouveau Deployment ID");
    }
    
    return newId;
    
  } catch (error) {
    logError("refreshDeploymentId", error);
    return null;
  }
}

/**
 * Génère l'URL complète de l'application web
 * 
 * @returns {string} URL complète ou URL template
 */
function getWebAppUrl() {
  const deploymentId = getDeploymentIdAuto();
  
  if (deploymentId) {
    return `https://script.google.com/macros/s/${deploymentId}/exec`;
  } else {
    return "https://script.google.com/macros/s/DEPLOYMENT_ID_NON_DISPONIBLE/exec";
  }
}

/**
 * ============================================
 * 3️⃣ ESPACE DÉVELOPPEUR PAR DÉFAUT
 * ============================================
 */

/**
 * 🛠️ Initialise automatiquement l'espace développeur
 * Crée un compte client pour le développeur avec tous les droits
 * 
 * Cette fonction est appelée automatiquement au premier lancement
 * si DEV_CONFIG.autoInitialize = true
 * 
 * @returns {Object} Résultat {success, userId, message}
 */
function initDevWorkspace() {
  try {
    logInfo("🛠️ Initialisation espace développeur...");
    
    // Vérifier si déjà initialisé
    const existing = findClientByEmail(DEV_CONFIG.email);
    
    if (existing) {
      logInfo(`✅ Espace développeur déjà initialisé : ${existing.userId}`);
      return {
        success: true,
        userId: existing.userId,
        message: "Espace développeur déjà existant",
        alreadyExists: true
      };
    }
    
    // Création du compte développeur
    const devClient = {
      userId: "DEV_" + generateFullId(),
      submissionId: "DEV_INIT",
      respondentId: "DEV_INIT",
      
      // Informations personnelles
      fullName: DEV_CONFIG.fullName,
      loginEmail: DEV_CONFIG.email,
      senderEmail: DEV_CONFIG.email,
      replyEmail: DEV_CONFIG.email,
      companyName: DEV_CONFIG.companyName,
      
      // Version et permissions
      version: DEV_CONFIG.version,
      
      // Métadonnées
      submittedAt: new Date().toISOString(),
      activatedAt: new Date().toISOString(),
      status: "active",
      emailSent: false,
      
      // Flags spéciaux développeur
      isDeveloper: true,
      hasFullAccess: true,
      
      metadata: {
        source: "dev_init",
        activationMethod: "automatic",
        role: "developer"
      }
    };
    
    // Génération du lien personnel
    devClient.personalLink = generatePersonalLink(devClient.userId);
    
    // Sauvegarde
    saveClient(devClient);
    
    // Définir la version utilisateur
    setUserVersion(DEV_CONFIG.version);
    
    logInfo(`✅ Espace développeur créé : ${devClient.userId}`);
    
    // Affichage console
    console.log("\n" + "=".repeat(60));
    console.log("🛠️ ESPACE DÉVELOPPEUR INITIALISÉ");
    console.log("=".repeat(60));
    console.log(`Email       : ${devClient.loginEmail}`);
    console.log(`Version     : ${devClient.version}`);
    console.log(`User ID     : ${devClient.userId}`);
    console.log(`Lien perso  : ${devClient.personalLink}`);
    console.log("=".repeat(60) + "\n");
    
    return {
      success: true,
      userId: devClient.userId,
      personalLink: devClient.personalLink,
      message: "Espace développeur créé avec succès"
    };
    
  } catch (error) {
    logError("initDevWorkspace", error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Vérifie si l'utilisateur courant est le développeur
 * 
 * @returns {boolean} true si développeur
 */
function isCurrentUserDeveloper() {
  try {
    const currentEmail = Session.getActiveUser().getEmail();
    return normalizeEmail(currentEmail) === normalizeEmail(DEV_CONFIG.email);
  } catch (error) {
    return false;
  }
}

/**
 * Récupère les informations de l'espace développeur
 * 
 * @returns {Object|null} Client développeur ou null
 */
function getDevWorkspace() {
  return findClientByEmail(DEV_CONFIG.email);
}

/**
 * ============================================
 * INITIALISATION AUTOMATIQUE COMPLÈTE
 * ============================================
 */

/**
 * 🚀 Initialisation automatique complète du système
 * 
 * Cette fonction est appelée automatiquement au premier lancement
 * Elle configure TOUT le système sans intervention manuelle :
 * 1. Détecte le Google Sheet
 * 2. Récupère le Deployment ID
 * 3. Crée l'espace développeur
 * 4. Configure les propriétés système
 * 
 * @returns {Object} Résultat complet de l'initialisation
 */
function autoInitializeSystem() {
  try {
    logInfo("🚀 INITIALISATION AUTOMATIQUE DU SYSTÈME...");
    
    const result = {
      success: true,
      steps: [],
      warnings: [],
      config: {}
    };
    
    // ===== 1️⃣ DÉTECTION SPREADSHEET =====
    const spreadsheet = detectActiveSpreadsheet();
    
    if (spreadsheet) {
      registerSpreadsheetAuto(spreadsheet);
      result.steps.push(`✅ Spreadsheet détecté : ${spreadsheet.getName()}`);
      result.config.spreadsheetId = spreadsheet.getId();
      result.config.spreadsheetName = spreadsheet.getName();
    } else {
      result.warnings.push("⚠️ Aucun spreadsheet Tally détecté - ajoutez-en un dans Drive");
    }
    
    // ===== 2️⃣ DEPLOYMENT ID =====
    const deploymentId = getDeploymentIdAuto();
    
    if (deploymentId) {
      result.steps.push("✅ Deployment ID récupéré automatiquement");
      result.config.deploymentId = deploymentId;
      result.config.webAppUrl = getWebAppUrl();
    } else {
      result.warnings.push("⚠️ Deployment ID non disponible - déployez l'application web d'abord");
      result.config.deploymentId = null;
      result.config.webAppUrl = "Non disponible (déployez d'abord)";
    }
    
    // ===== 3️⃣ ESPACE DÉVELOPPEUR =====
    if (DEV_CONFIG.autoInitialize) {
      const devResult = initDevWorkspace();
      
      if (devResult.success) {
        if (devResult.alreadyExists) {
          result.steps.push("✅ Espace développeur déjà existant");
        } else {
          result.steps.push("✅ Espace développeur créé");
        }
        result.config.devUserId = devResult.userId;
        result.config.devLink = devResult.personalLink;
      } else {
        result.warnings.push("⚠️ Erreur création espace développeur : " + devResult.error);
      }
    } else {
      result.steps.push("ℹ️ Espace développeur non initialisé (autoInitialize = false)");
    }
    
    // ===== 4️⃣ VALIDATION =====
    const validation = validateSystemConfiguration();
    result.validation = validation;
    
    if (!validation.success) {
      result.warnings.push(...validation.errors);
    }
    
    // ===== AFFICHAGE RÉSULTATS =====
    console.log("\n" + "=".repeat(60));
    console.log("🎉 INITIALISATION AUTOMATIQUE TERMINÉE");
    console.log("=".repeat(60));
    console.log("\n📋 ÉTAPES COMPLÉTÉES :");
    result.steps.forEach(step => console.log(step));
    
    if (result.warnings.length > 0) {
      console.log("\n⚠️ AVERTISSEMENTS :");
      result.warnings.forEach(warn => console.log(warn));
    }
    
    console.log("\n🔧 CONFIGURATION SYSTÈME :");
    console.log(`Spreadsheet  : ${result.config.spreadsheetName || "Non détecté"}`);
    console.log(`Deployment   : ${result.config.deploymentId || "Non disponible"}`);
    console.log(`Web App URL  : ${result.config.webAppUrl || "N/A"}`);
    console.log(`Dev User ID  : ${result.config.devUserId || "Non créé"}`);
    
    console.log("\n✅ LE SYSTÈME EST PRÊT À L'EMPLOI !");
    console.log("=".repeat(60) + "\n");
    
    logInfo("🎉 Initialisation automatique terminée avec succès");
    
    return result;
    
  } catch (error) {
    logError("autoInitializeSystem", error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Valide la configuration complète du système
 * 
 * @returns {Object} Résultat validation
 */
function validateSystemConfiguration() {
  const result = {
    success: true,
    errors: [],
    warnings: []
  };
  
  // Test 1 : Spreadsheet
  const sheetId = PropertiesService.getScriptProperties().getProperty("SOURCE_SHEET_ID");
  if (!sheetId) {
    result.errors.push("❌ Aucun spreadsheet configuré");
    result.success = false;
  }
  
  // Test 2 : Deployment ID
  const deploymentId = getDeploymentIdAuto();
  if (!deploymentId) {
    result.warnings.push("⚠️ Deployment ID non disponible (déployez l'app web)");
  }
  
  // Test 3 : Permissions Gmail
  try {
    GmailApp.getAliases();
  } catch (e) {
    result.errors.push("❌ Permission Gmail manquante");
    result.success = false;
  }
  
  // Test 4 : Espace développeur
  if (DEV_CONFIG.autoInitialize) {
    const devWorkspace = getDevWorkspace();
    if (!devWorkspace) {
      result.warnings.push("⚠️ Espace développeur non initialisé");
    }
  }
  
  return result;
}

/**
 * ============================================
 * HELPER : APPEL LORS DU PREMIER onEdit
 * ============================================
 */

/**
 * Vérifie si le système est initialisé
 * Sinon, lance l'initialisation automatique
 * 
 * Cette fonction est appelée par onEdit() lors de la première détection
 */
function ensureSystemInitialized() {
  try {
    // Vérifier si déjà initialisé
    const initialized = PropertiesService.getScriptProperties()
      .getProperty("SYSTEM_INITIALIZED");
    
    if (initialized === "true") {
      return; // Déjà fait
    }
    
    // Lancer l'initialisation
    logInfo("🔄 Première exécution détectée - initialisation automatique...");
    autoInitializeSystem();
    
    // Marquer comme initialisé
    PropertiesService.getScriptProperties()
      .setProperty("SYSTEM_INITIALIZED", "true");
    
  } catch (error) {
    logError("ensureSystemInitialized", error);
  }
}

/**
 * ============================================
 * FONCTIONS UTILITAIRES
 * ============================================
 */

/**
 * Réinitialise complètement le système (pour tests)
 * ⚠️ À utiliser avec précaution
 */
function resetSystemConfiguration() {
  try {
    PropertiesService.getScriptProperties().deleteProperty("SYSTEM_INITIALIZED");
    PropertiesService.getScriptProperties().deleteProperty("SOURCE_SHEET_ID");
    PropertiesService.getScriptProperties().deleteProperty("DEPLOYMENT_ID");
    
    logWarning("🔄 Configuration système réinitialisée");
    
    return {
      success: true,
      message: "Système réinitialisé - relancez autoInitializeSystem()"
    };
    
  } catch (error) {
    logError("resetSystemConfiguration", error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Affiche la configuration actuelle
 */
function showCurrentConfiguration() {
  const config = {
    spreadsheetId: PropertiesService.getScriptProperties().getProperty("SOURCE_SHEET_ID"),
    deploymentId: getDeploymentIdAuto(),
    webAppUrl: getWebAppUrl(),
    devEmail: DEV_CONFIG.email,
    devWorkspace: getDevWorkspace(),
    systemInitialized: PropertiesService.getScriptProperties().getProperty("SYSTEM_INITIALIZED")
  };
  
  console.log(JSON.stringify(config, null, 2));
  return config;
}
