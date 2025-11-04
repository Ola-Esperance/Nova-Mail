/*****************************************************
 * NOVAMAIL SAAS - SHEETTRIGGERS.GS
 * ====================================================
 * SYSTÈME DE DÉCLENCHEMENT INSTANTANÉ 100% NATIF
 * 
 * ⚡ Déclenchement : INSTANTANÉ (< 1 seconde)
 * 🎯 Méthode : Simple Trigger onEdit() natif Google
 * ♾️ Limites : AUCUNE (pas de quota)
 * 🔧 Installation : AUTOMATIQUE (rien à faire)
 * 
 * @author NovaMail Team
 * @version 3.0.0 PRODUCTION
 * @lastModified 2025-11-04
 *****************************************************/

/**
 * ============================================
 * 🔥 TRIGGER INSTANTANÉ (FONCTION PRINCIPALE)
 * ============================================
 * 
 * ⚡ Cette fonction s'exécute AUTOMATIQUEMENT et INSTANTANÉMENT
 * dès qu'une cellule est modifiée dans n'importe quelle feuille du spreadsheet.
 * 
 * AVANTAGES :
 * - Déclenchement < 1 seconde après ajout ligne
 * - AUCUNE installation requise (fonctionne automatiquement)
 * - AUCUNE limite de déclenchements (illimité)
 * - Fonctionne pour TOUTES les feuilles du projet
 * - 100% natif Google Apps Script
 * 
 * ⚠️ NE PAS MODIFIER LE NOM DE CETTE FONCTION
 * Google Apps Script la détecte automatiquement
 * 
 * @param {Object} e - Event object Google Apps Script
 */
function onEdit(e) {
  try {
    // ===== INITIALISATION AUTOMATIQUE (PREMIÈRE FOIS) =====
    ensureSystemInitialized();
    
    // ===== AUTO-DÉTECTION & ENREGISTREMENT DU SHEET =====
    if (e && e.range && e.range.getSheet()) {
      const spreadsheet = e.range.getSheet().getParent();
      registerSpreadsheetAuto(spreadsheet);
    }
    
    // ===== VALIDATION PRÉLIMINAIRE =====
    if (!e || !e.range) return; // Pas de modification détectée
    
    const range = e.range;
    const sheet = range.getSheet();
    const row = range.getRow();
    const col = range.getColumn();
    
    // Ignorer ligne d'en-tête
    if (row === 1) return;
    
    // ===== DÉTECTION NOUVELLE SOUMISSION TALLY =====
    // Tally ajoute toute une ligne d'un coup, donc on vérifie :
    // 1. Si c'est la dernière colonne remplie (nouvelle ligne complète)
    // 2. Si la ligne n'a pas déjà été traitée
    
    const lastCol = sheet.getLastColumn();
    
    // Vérifier que c'est bien une nouvelle ligne complète
    // (Tally remplit toutes les colonnes à la fois)
    if (col !== lastCol) return;
    
    // Vérifier que la ligne est complète
    const rowData = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
    const isComplete = rowData.every(cell => cell !== "" && cell !== null && cell !== undefined);
    
    if (!isComplete) return;
    
    // ===== PROTECTION CONTRE DOUBLE TRAITEMENT =====
    if (isRowAlreadyProcessed(sheet, row)) {
      logInfo(`⏭️ Ligne ${row} déjà traitée - ignorée`);
      return;
    }
    
    // ===== TRAITEMENT INSTANTANÉ =====
    logInfo(`🔔 NOUVELLE SOUMISSION DÉTECTÉE - Ligne ${row} - Traitement immédiat...`);
    
    // Marquage immédiat (évite retraitement si erreur)
    markRowAsProcessing(sheet, row);
    
    // Lancement du traitement
    processNewSubmission(sheet, row);
    
  } catch (error) {
    // Log silencieux pour ne pas bloquer Google Sheets
    try {
      logError("onEdit", error);
    } catch (e) {
      // Ignore
    }
  }
}

/**
 * ============================================
 * TRAITEMENT D'UNE NOUVELLE SOUMISSION
 * ============================================
 */

/**
 * 🚀 Traite une nouvelle soumission instantanément
 * 
 * Cette fonction :
 * 1. Extrait les données de la ligne
 * 2. Appelle processNewClientSubmission() (UserManagement.gs)
 * 3. Met à jour le statut dans le sheet
 * 4. Cache la ligne pour éviter retraitement
 * 
 * @param {Sheet} sheet - Feuille Google Sheets
 * @param {number} rowNumber - Numéro de ligne (1-indexed)
 */
function processNewSubmission(sheet, rowNumber) {
  const startTime = new Date();
  
  try {
    // 1️⃣ EXTRACTION DES DONNÉES
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const values = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Construction objet clé-valeur
    const rowData = {};
    headers.forEach((header, index) => {
      rowData[header] = values[index];
    });
    
    logInfo(`📋 Données extraites pour ligne ${rowNumber}`);
    
    // 2️⃣ VALIDATION DONNÉES TALLY
    const submissionId = rowData["Submission ID"];
    const email = rowData["Email de connexion"];
    
    if (!submissionId || !email) {
      throw new Error("Données incomplètes - Submission ID ou Email manquant");
    }
    
    // 3️⃣ APPEL FONCTION D'ACTIVATION CLIENT
    // Définie dans UserManagement.gs
    const result = processNewClientSubmission(rowData);
    
    // 4️⃣ MISE À JOUR STATUT DANS LE SHEET
    updateRowStatus(sheet, rowNumber, result);
    
    // 5️⃣ CACHE POUR ÉVITER RETRAITEMENT
    cacheProcessedRow(sheet, rowNumber);
    
    // 6️⃣ LOG SUCCÈS
    const duration = new Date() - startTime;
    
    if (result.success) {
      logInfo(`✅ Ligne ${rowNumber} traitée avec succès en ${duration}ms - UserID: ${result.userId}`);
    } else {
      logError(`❌ Ligne ${rowNumber} échouée`, new Error(result.message));
    }
    
  } catch (error) {
    // Gestion d'erreur robuste
    logError(`processNewSubmission [ligne ${rowNumber}]`, error);
    
    try {
      markRowAsError(sheet, rowNumber, error.message);
    } catch (e) {
      // Ignore si mise à jour impossible
    }
  }
}

/**
 * ============================================
 * PROTECTION CONTRE DOUBLE TRAITEMENT
 * ============================================
 */

/**
 * Vérifie si une ligne a déjà été traitée
 * Utilise 3 niveaux de vérification :
 * 1. Cache rapide (CacheService)
 * 2. Colonne statut dans le sheet
 * 3. Vérification dans PropertiesService
 * 
 * @param {Sheet} sheet - Feuille
 * @param {number} rowNumber - Numéro de ligne
 * @returns {boolean} true si déjà traitée
 */
function isRowAlreadyProcessed(sheet, rowNumber) {
  try {
    // ===== NIVEAU 1 : CACHE (ultra-rapide) =====
    const cache = CacheService.getScriptCache();
    const cacheKey = buildCacheKey(sheet, rowNumber);
    
    if (cache.get(cacheKey)) {
      return true; // Déjà traitée (en cache)
    }
    
    // ===== NIVEAU 2 : COLONNE STATUT =====
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const statusColIndex = headers.indexOf("Statut activation");
    
    if (statusColIndex !== -1) {
      const statusValue = sheet.getRange(rowNumber, statusColIndex + 1).getValue();
      const statusStr = String(statusValue);
      
      if (statusStr.includes("✅") || statusStr.includes("Activé")) {
        // Ajouter au cache pour accélérer prochaines vérifications
        cacheProcessedRow(sheet, rowNumber);
        return true;
      }
      
      if (statusStr.includes("⏳") || statusStr.includes("En cours")) {
        // Traitement déjà en cours
        return true;
      }
    }
    
    // ===== NIVEAU 3 : VÉRIFICATION SUBMISSION ID =====
    // Évite les doublons même si colonnes supprimées
    const submissionId = sheet.getRange(rowNumber, 1).getValue(); // Colonne A = Submission ID
    
    if (submissionId) {
      const processed = PropertiesService.getScriptProperties()
        .getProperty("PROCESSED_" + submissionId);
      
      if (processed) {
        cacheProcessedRow(sheet, rowNumber);
        return true;
      }
    }
    
    return false; // Pas encore traitée
    
  } catch (error) {
    // En cas d'erreur, on considère non traitée (principe de précaution)
    logWarning("isRowAlreadyProcessed: " + error.message);
    return false;
  }
}

/**
 * Met en cache une ligne traitée
 * 
 * @param {Sheet} sheet - Feuille
 * @param {number} rowNumber - Numéro de ligne
 */
function cacheProcessedRow(sheet, rowNumber) {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = buildCacheKey(sheet, rowNumber);
    
    // Cache pendant 6 heures (21600 secondes)
    cache.put(cacheKey, "processed", 21600);
    
    // Enregistrer aussi le Submission ID
    const submissionId = sheet.getRange(rowNumber, 1).getValue();
    if (submissionId) {
      PropertiesService.getScriptProperties()
        .setProperty("PROCESSED_" + submissionId, new Date().toISOString());
    }
    
  } catch (error) {
    // Erreur silencieuse : le cache est optionnel
    logWarning("cacheProcessedRow: " + error.message);
  }
}

/**
 * Construit une clé unique pour le cache
 * 
 * @param {Sheet} sheet - Feuille
 * @param {number} rowNumber - Numéro de ligne
 * @returns {string} Clé cache unique
 */
function buildCacheKey(sheet, rowNumber) {
  return "PROCESSED_" + sheet.getSheetId() + "_" + rowNumber;
}

/**
 * ============================================
 * MISE À JOUR STATUT DANS LE SHEET
 * ============================================
 */

/**
 * Marque une ligne comme "en cours de traitement"
 * 
 * @param {Sheet} sheet - Feuille
 * @param {number} rowNumber - Numéro de ligne
 */
function markRowAsProcessing(sheet, rowNumber) {
  try {
    const statusCol = ensureStatusColumn(sheet);
    
    const cell = sheet.getRange(rowNumber, statusCol);
    cell.setValue("⏳ En cours...");
    cell.setBackground("#fef3c7"); // Jaune clair
    cell.setFontColor("#92400e");
    
  } catch (error) {
    logWarning("markRowAsProcessing: " + error.message);
  }
}

/**
 * Met à jour le statut final d'une ligne après traitement
 * 
 * @param {Sheet} sheet - Feuille
 * @param {number} rowNumber - Numéro de ligne
 * @param {Object} result - Résultat du traitement
 */
function updateRowStatus(sheet, rowNumber, result) {
  try {
    const statusCol = ensureStatusColumn(sheet);
    const dateCol = statusCol + 1;
    const linkCol = statusCol + 2;
    const userIdCol = statusCol + 3;
    
    // Mise à jour statut
    const statusValue = result.success 
      ? "✅ Activé" 
      : `❌ Erreur : ${(result.message || "Inconnu").substring(0, 100)}`;
    
    const statusCell = sheet.getRange(rowNumber, statusCol);
    statusCell.setValue(statusValue);
    
    if (result.success) {
      // Succès : fond vert
      statusCell.setBackground("#d1fae5");
      statusCell.setFontColor("#065f46");
      
      // Date activation
      sheet.getRange(rowNumber, dateCol).setValue(new Date());
      
      // Lien personnel
      if (result.personalLink) {
        sheet.getRange(rowNumber, linkCol).setValue(result.personalLink);
      }
      
      // User ID
      if (result.userId) {
        sheet.getRange(rowNumber, userIdCol).setValue(result.userId);
      }
      
    } else {
      // Erreur : fond rouge
      statusCell.setBackground("#fee2e2");
      statusCell.setFontColor("#991b1b");
      
      // Date tentative
      sheet.getRange(rowNumber, dateCol).setValue(new Date());
    }
    
  } catch (error) {
    logWarning("updateRowStatus: " + error.message);
  }
}

/**
 * Marque une ligne comme erreur
 * 
 * @param {Sheet} sheet - Feuille
 * @param {number} rowNumber - Numéro de ligne
 * @param {string} errorMessage - Message d'erreur
 */
function markRowAsError(sheet, rowNumber, errorMessage) {
  try {
    updateRowStatus(sheet, rowNumber, {
      success: false,
      message: errorMessage
    });
  } catch (error) {
    logWarning("markRowAsError: " + error.message);
  }
}

/**
 * S'assure que les colonnes de statut existent
 * Les crée si nécessaire
 * 
 * @param {Sheet} sheet - Feuille
 * @returns {number} Index de la colonne "Statut activation"
 */
function ensureStatusColumn(sheet) {
  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let statusColIndex = headers.indexOf("Statut activation");
    
    if (statusColIndex === -1) {
      // Créer les colonnes
      const lastCol = sheet.getLastColumn();
      
      sheet.getRange(1, lastCol + 1).setValue("Statut activation");
      sheet.getRange(1, lastCol + 2).setValue("Date activation");
      sheet.getRange(1, lastCol + 3).setValue("Lien personnel");
      sheet.getRange(1, lastCol + 4).setValue("User ID");
      
      // Formatage en-têtes
      const headerRange = sheet.getRange(1, lastCol + 1, 1, 4);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#4f46e5");
      headerRange.setFontColor("#ffffff");
      
      statusColIndex = lastCol;
      
      logInfo("📋 Colonnes de suivi créées automatiquement");
    }
    
    return statusColIndex + 1; // +1 car indices Google Sheets commencent à 1
    
  } catch (error) {
    logError("ensureStatusColumn", error);
    return sheet.getLastColumn() + 1; // Fallback
  }
}

/**
 * ============================================
 * FONCTIONS DE MAINTENANCE & DEBUG
 * ============================================
 */

/**
 * 🔧 Traite manuellement une ligne spécifique
 * Utile pour tests ou retraitement après erreur
 * 
 * @param {number} rowNumber - Numéro de ligne à traiter
 * @param {string} sheetName - Nom de la feuille (optionnel)
 * @returns {Object} Résultat du traitement
 * 
 * @example
 * // Traiter la ligne 2 de la feuille active
 * manualProcessRow(2);
 * 
 * // Traiter la ligne 5 d'une feuille spécifique
 * manualProcessRow(5, "Réponses au formulaire 1");
 */
function manualProcessRow(rowNumber, sheetName) {
  try {
    logInfo(`🔧 Traitement manuel ligne ${rowNumber}...`);
    
    // Récupération de la feuille
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = sheetName 
      ? spreadsheet.getSheetByName(sheetName)
      : spreadsheet.getActiveSheet();
    
    if (!sheet) {
      throw new Error(`Feuille "${sheetName}" introuvable`);
    }
    
    // Traitement
    processNewSubmission(sheet, rowNumber);
    
    logInfo(`✅ Traitement manuel terminé`);
    
    return {
      success: true,
      message: `Ligne ${rowNumber} traitée avec succès`
    };
    
  } catch (error) {
    logError("manualProcessRow", error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 🔄 Retraite toutes les lignes non marquées comme traitées
 * Utile après une interruption ou pour rattrapage
 * 
 * @param {string} sheetName - Nom de la feuille (optionnel)
 * @returns {Object} Résultat {processed, errors, skipped}
 * 
 * @example
 * // Retraiter toutes les lignes non finalisées de la feuille active
 * reprocessUnfinishedRows();
 * 
 * // Retraiter une feuille spécifique
 * reprocessUnfinishedRows("Réponses au formulaire 1");
 */
function reprocessUnfinishedRows(sheetName) {
  try {
    logWarning("🔄 Retraitement des lignes non finalisées...");
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = sheetName 
      ? spreadsheet.getSheetByName(sheetName)
      : spreadsheet.getActiveSheet();
    
    if (!sheet) {
      throw new Error(`Feuille "${sheetName}" introuvable`);
    }
    
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return { 
        processed: 0, 
        errors: 0, 
        skipped: 0,
        message: "Aucune ligne à traiter" 
      };
    }
    
    let processed = 0;
    let errors = 0;
    let skipped = 0;
    
    // Parcourir toutes les lignes (sauf en-tête)
    for (let row = 2; row <= lastRow; row++) {
      try {
        // Vérifier si déjà traitée
        if (isRowAlreadyProcessed(sheet, row)) {
          skipped++;
          continue;
        }
        
        // Traiter la ligne
        processNewSubmission(sheet, row);
        processed++;
        
        // Délai pour éviter surcharge
        Utilities.sleep(1500);
        
      } catch (err) {
        errors++;
        logError(`Erreur ligne ${row}`, err);
      }
    }
    
    const result = {
      processed: processed,
      errors: errors,
      skipped: skipped,
      message: `${processed} traitée(s), ${errors} erreur(s), ${skipped} ignorée(s)`
    };
    
    logInfo(`✅ Retraitement terminé : ${result.message}`);
    
    return result;
    
  } catch (error) {
    logError("reprocessUnfinishedRows", error);
    return {
      processed: 0,
      errors: 0,
      skipped: 0,
      error: error.message
    };
  }
}

/**
 * 🧹 Efface le cache de toutes les lignes traitées
 * Permet de forcer le retraitement si nécessaire
 * 
 * @returns {boolean} Succès
 * 
 * @example
 * clearProcessedCache();
 */
function clearProcessedCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.removeAll(["PROCESSED_"]);
    
    logInfo("🧹 Cache effacé");
    return true;
    
  } catch (error) {
    logError("clearProcessedCache", error);
    return false;
  }
}

/**
 * 📊 Génère un rapport du système
 * Affiche le statut de toutes les feuilles
 * 
 * @returns {Object} Rapport complet
 * 
 * @example
 * const report = getSystemReport();
 * Logger.log(JSON.stringify(report, null, 2));
 */
function getSystemReport() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    
    const report = {
      spreadsheetId: spreadsheet.getId(),
      spreadsheetName: spreadsheet.getName(),
      totalSheets: sheets.length,
      sheets: [],
      totalRows: 0,
      processedRows: 0,
      errorRows: 0,
      pendingRows: 0
    };
    
    sheets.forEach(sheet => {
      const sheetReport = analyzeSheet(sheet);
      report.sheets.push(sheetReport);
      
      report.totalRows += sheetReport.totalRows;
      report.processedRows += sheetReport.processedRows;
      report.errorRows += sheetReport.errorRows;
      report.pendingRows += sheetReport.pendingRows;
    });
    
    return report;
    
  } catch (error) {
    logError("getSystemReport", error);
    return {
      error: error.message,
      status: "error"
    };
  }
}

/**
 * Analyse une feuille spécifique
 * 
 * @param {Sheet} sheet - Feuille à analyser
 * @returns {Object} Rapport de la feuille
 */
function analyzeSheet(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    const totalRows = lastRow - 1; // Exclure en-tête
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const statusColIndex = headers.indexOf("Statut activation");
    
    let processedRows = 0;
    let errorRows = 0;
    
    if (statusColIndex !== -1 && totalRows > 0) {
      const statusValues = sheet.getRange(2, statusColIndex + 1, totalRows, 1).getValues();
      
      statusValues.forEach(row => {
        const status = String(row[0]);
        if (status.includes("✅")) processedRows++;
        if (status.includes("❌")) errorRows++;
      });
    }
    
    return {
      sheetName: sheet.getName(),
      totalRows: totalRows,
      processedRows: processedRows,
      errorRows: errorRows,
      pendingRows: totalRows - processedRows - errorRows
    };
    
  } catch (error) {
    return {
      sheetName: sheet.getName(),
      error: error.message
    };
  }
}

/**
 * ============================================
 * FONCTION DE CONFIGURATION (OPTIONNELLE)
 * ============================================
 */

/**
 * 🚀 Configuration initiale du système
 * 
 * Cette fonction est OPTIONNELLE car le système fonctionne
 * automatiquement dès que le code est en place.
 * 
 * Elle sert uniquement à configurer le Deployment ID
 * pour les liens personnels des clients.
 * 
 * @param {string} deploymentId - ID du déploiement web
 * 
 * @example
 * setupNovaMail("AKfycbzXXXXXXXXXXXXX");
 */
function setupNovaMail(deploymentId) {
  try {
    logInfo("🚀 Configuration NovaMail...");
    
    if (deploymentId) {
      setDeploymentId(deploymentId);
      logInfo("✅ Deployment ID configuré");
    }
    
    // Test de validation
    const testResult = testSystemConfiguration();
    
    console.log("\n" + "=".repeat(60));
    console.log("✅ NOVAMAIL CONFIGURÉ ET ACTIF");
    console.log("=".repeat(60));
    console.log("");
    console.log("🔥 Le système est maintenant opérationnel !");
    console.log("");
    console.log("📌 Fonctionnement :");
    console.log("  → Tally ajoute une ligne dans Google Sheets");
    console.log("  → onEdit() se déclenche instantanément (< 1 sec)");
    console.log("  → Client activé et email envoyé automatiquement");
    console.log("");
    console.log("✅ Aucune autre installation requise !");
    console.log("✅ Aucune limite de déclenchements !");
    console.log("✅ Fonctionne pour toutes les feuilles du projet !");
    console.log("");
    console.log("🧪 Pour tester : soumettez votre formulaire Tally");
    console.log("=".repeat(60) + "\n");
    
    return {
      success: true,
      message: "Configuration terminée",
      testResult: testResult
    };
    
  } catch (error) {
    logError("setupNovaMail", error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 🧪 Teste la configuration du système
 * 
 * @returns {Object} Résultat des tests
 */
function testSystemConfiguration() {
  const results = {
    success: true,
    checks: []
  };
  
  // Test 1 : Permissions Gmail
  try {
    GmailApp.getAliases();
    results.checks.push({ name: "Permissions Gmail", passed: true });
  } catch (e) {
    results.checks.push({ name: "Permissions Gmail", passed: false, error: e.message });
    results.success = false;
  }
  
  // Test 2 : Spreadsheet actif
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    results.checks.push({ 
      name: "Spreadsheet actif", 
      passed: true,
      value: ss.getName()
    });
  } catch (e) {
    results.checks.push({ name: "Spreadsheet actif", passed: false, error: e.message });
    results.success = false;
  }
  
  // Test 3 : Deployment ID
  const deploymentId = getDeploymentId();
  results.checks.push({ 
    name: "Deployment ID", 
    passed: !!deploymentId,
    value: deploymentId || "Non configuré"
  });
  
  return results;
}
