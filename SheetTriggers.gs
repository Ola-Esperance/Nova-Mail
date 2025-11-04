/*****************************************************
 * NOVAMAIL SAAS - SHEETTRIGGERS.GS (VERSION FIXÉE)
 * ====================================================
 * ✅ FIX : Utilise Sheet ID direct au lieu du parcours Drive
 * ✅ FIX : Logs détaillés à chaque étape
 * ✅ FIX : Gestion robuste des erreurs
 * 
 * @version 3.1.0 FIXED
 * @lastModified 2025-11-04
 *****************************************************/

/**
 * ============================================
 * 🔥 TRIGGER INSTANTANÉ (FONCTION PRINCIPALE)
 * ============================================
 */

function onEdit(e) {
  try {
    // ===== INITIALISATION AUTOMATIQUE =====
    ensureSystemInitialized();
    
    // ===== AUTO-DÉTECTION & ENREGISTREMENT DU SHEET =====
    if (e && e.range && e.range.getSheet()) {
      const spreadsheet = e.range.getSheet().getParent();
      const sheetId = spreadsheet.getId();
      
      // Enregistrement automatique si pas déjà fait
      const currentId = getSourceSheetId();
      if (!currentId || currentId !== sheetId) {
        setSourceSheetId(sheetId);
        logInfo(`📌 Sheet auto-enregistré : ${sheetId} (${spreadsheet.getName()})`);
      }
    }
    
    // ===== VALIDATION PRÉLIMINAIRE =====
    if (!e || !e.range) return;
    
    const range = e.range;
    const sheet = range.getSheet();
    const row = range.getRow();
    const col = range.getColumn();
    
    // Ignorer ligne d'en-tête
    if (row === 1) return;
    
    // ===== DÉTECTION NOUVELLE SOUMISSION TALLY =====
    const lastCol = sheet.getLastColumn();
    
    if (col !== lastCol) return;
    
    // Vérifier ligne complète
    const rowData = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
    const isComplete = rowData.every(cell => cell !== "" && cell !== null && cell !== undefined);
    
    if (!isComplete) return;
    
    // ===== PROTECTION CONTRE DOUBLE TRAITEMENT =====
    if (isRowAlreadyProcessed(sheet, row)) {
      logInfo(`⏭️ Ligne ${row} déjà traitée - ignorée`);
      return;
    }
    
    // ===== TRAITEMENT INSTANTANÉ =====
    logInfo(`🔔 NOUVELLE SOUMISSION DÉTECTÉE - Ligne ${row}`);
    
    markRowAsProcessing(sheet, row);
    processNewSubmission(sheet, row);
    
  } catch (error) {
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

function processNewSubmission(sheet, rowNumber) {
  const startTime = new Date();
  const logId = generateShortId();
  
  try {
    logInfo(`[${logId}] 🚀 Traitement ligne ${rowNumber} démarré`);
    
    // 1️⃣ EXTRACTION DES DONNÉES
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const values = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const rowData = {};
    headers.forEach((header, index) => {
      rowData[header] = values[index];
    });
    
    logInfo(`[${logId}] ✅ Données extraites : ${Object.keys(rowData).length} colonnes`);
    
    // 2️⃣ VALIDATION DONNÉES TALLY
    const submissionId = rowData["Submission ID"];
    const email = rowData["Email de connexion"];
    
    if (!submissionId || !email) {
      throw new Error("Données incomplètes - Submission ID ou Email manquant");
    }
    
    logInfo(`[${logId}] 📧 Email détecté : ${email}`);
    
    // 3️⃣ APPEL FONCTION D'ACTIVATION CLIENT
    logInfo(`[${logId}] 🔄 Appel processNewClientSubmission...`);
    
    const result = processNewClientSubmission(rowData);
    
    logInfo(`[${logId}] ✅ Résultat activation : ${JSON.stringify(result)}`);
    
    // 4️⃣ MISE À JOUR STATUT
    updateRowStatus(sheet, rowNumber, result);
    
    // 5️⃣ CACHE
    cacheProcessedRow(sheet, rowNumber);
    
    // 6️⃣ LOG SUCCÈS
    const duration = new Date() - startTime;
    
    if (result.success) {
      logInfo(`[${logId}] ✅ Ligne ${rowNumber} traitée en ${duration}ms - UserID: ${result.userId}`);
    } else {
      logError(`[${logId}] ❌ Ligne ${rowNumber} échouée`, new Error(result.message));
      
      // Notification développeur si échec
      notifyDeveloperOfActivationFailure(rowData, result, logId);
    }
    
  } catch (error) {
    logError(`processNewSubmission [ligne ${rowNumber}]`, error);
    
    try {
      markRowAsError(sheet, rowNumber, error.message);
      
      // Notification développeur
      notifyDeveloperOfActivationFailure({
        rowNumber: rowNumber,
        sheetName: sheet.getName()
      }, {
        success: false,
        message: error.message
      }, logId || "unknown");
      
    } catch (e) {
      logWarning("Impossible de mettre à jour le statut d'erreur");
    }
  }
}

/**
 * ============================================
 * NOTIFICATION DÉVELOPPEUR EN CAS D'ÉCHEC
 * ============================================
 */

function notifyDeveloperOfActivationFailure(rowData, result, logId) {
  try {
    const devEmail = DEV_CONFIG.email || getDefaultSenderEmail();
    
    if (!devEmail || !isValidEmail(devEmail)) {
      logWarning("Email développeur invalide - notification impossible");
      return;
    }
    
    const subject = `🚨 NovaMail - Échec activation client [${logId}]`;
    
    const htmlBody = `
      <h2>⚠️ Échec d'activation client</h2>
      <p>Une soumission Tally n'a pas pu être traitée correctement.</p>
      
      <h3>Informations de l'erreur</h3>
      <ul>
        <li><strong>Date/Heure:</strong> ${new Date().toLocaleString("fr-FR")}</li>
        <li><strong>Log ID:</strong> ${logId}</li>
        <li><strong>Message:</strong> ${result.message || "Erreur inconnue"}</li>
      </ul>
      
      <h3>Données de la soumission</h3>
      <ul>
        ${Object.entries(rowData).map(([key, value]) => 
          `<li><strong>${key}:</strong> ${value}</li>`
        ).join('')}
      </ul>
      
      <h3>Actions recommandées</h3>
      <ol>
        <li>Vérifier les données dans le Google Sheet</li>
        <li>Consulter les logs Apps Script (Vue → Exécutions)</li>
        <li>Vérifier que tous les champs requis sont remplis</li>
        <li>Exécuter manuellement : <code>manualProcessRow(rowNumber)</code></li>
      </ol>
      
      <hr>
      <p style="font-size:12px; color:#666;">
        NovaMail Error Reporter<br>
        Projet: ${ScriptApp.getScriptId()}
      </p>
    `;
    
    GmailApp.sendEmail(
      devEmail,
      subject,
      stripHtml(htmlBody),
      {
        htmlBody: htmlBody,
        name: "NovaMail Error Reporter"
      }
    );
    
    logInfo(`🚨 Notification échec envoyée à ${devEmail}`);
    
  } catch (error) {
    logError("notifyDeveloperOfActivationFailure", error);
  }
}

/**
 * ============================================
 * PROTECTION CONTRE DOUBLE TRAITEMENT
 * ============================================
 */

function isRowAlreadyProcessed(sheet, rowNumber) {
  try {
    // NIVEAU 1 : CACHE
    const cache = CacheService.getScriptCache();
    const cacheKey = buildCacheKey(sheet, rowNumber);
    
    if (cache.get(cacheKey)) {
      return true;
    }
    
    // NIVEAU 2 : COLONNE STATUT
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const statusColIndex = headers.indexOf("Statut activation");
    
    if (statusColIndex !== -1) {
      const statusValue = sheet.getRange(rowNumber, statusColIndex + 1).getValue();
      const statusStr = String(statusValue);
      
      if (statusStr.includes("✅") || statusStr.includes("Activé")) {
        cacheProcessedRow(sheet, rowNumber);
        return true;
      }
      
      if (statusStr.includes("⏳") || statusStr.includes("En cours")) {
        return true;
      }
    }
    
    // NIVEAU 3 : SUBMISSION ID
    const submissionId = sheet.getRange(rowNumber, 1).getValue();
    
    if (submissionId) {
      const processed = PropertiesService.getScriptProperties()
        .getProperty("PROCESSED_" + submissionId);
      
      if (processed) {
        cacheProcessedRow(sheet, rowNumber);
        return true;
      }
    }
    
    return false;
    
  } catch (error) {
    logWarning("isRowAlreadyProcessed: " + error.message);
    return false;
  }
}

function cacheProcessedRow(sheet, rowNumber) {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = buildCacheKey(sheet, rowNumber);
    
    cache.put(cacheKey, "processed", 21600);
    
    const submissionId = sheet.getRange(rowNumber, 1).getValue();
    if (submissionId) {
      PropertiesService.getScriptProperties()
        .setProperty("PROCESSED_" + submissionId, new Date().toISOString());
    }
    
  } catch (error) {
    logWarning("cacheProcessedRow: " + error.message);
  }
}

function buildCacheKey(sheet, rowNumber) {
  return "PROCESSED_" + sheet.getSheetId() + "_" + rowNumber;
}

/**
 * ============================================
 * MISE À JOUR STATUT DANS LE SHEET
 * ============================================
 */

function markRowAsProcessing(sheet, rowNumber) {
  try {
    const statusCol = ensureStatusColumn(sheet);
    
    const cell = sheet.getRange(rowNumber, statusCol);
    cell.setValue("⏳ En cours...");
    cell.setBackground("#fef3c7");
    cell.setFontColor("#92400e");
    
  } catch (error) {
    logWarning("markRowAsProcessing: " + error.message);
  }
}

function updateRowStatus(sheet, rowNumber, result) {
  try {
    const statusCol = ensureStatusColumn(sheet);
    const dateCol = statusCol + 1;
    const linkCol = statusCol + 2;
    const userIdCol = statusCol + 3;
    
    const statusValue = result.success 
      ? "✅ Activé" 
      : `❌ Erreur : ${(result.message || "Inconnu").substring(0, 100)}`;
    
    const statusCell = sheet.getRange(rowNumber, statusCol);
    statusCell.setValue(statusValue);
    
    if (result.success) {
      statusCell.setBackground("#d1fae5");
      statusCell.setFontColor("#065f46");
      
      sheet.getRange(rowNumber, dateCol).setValue(new Date());
      
      if (result.personalLink) {
        sheet.getRange(rowNumber, linkCol).setValue(result.personalLink);
      }
      
      if (result.userId) {
        sheet.getRange(rowNumber, userIdCol).setValue(result.userId);
      }
      
    } else {
      statusCell.setBackground("#fee2e2");
      statusCell.setFontColor("#991b1b");
      
      sheet.getRange(rowNumber, dateCol).setValue(new Date());
    }
    
  } catch (error) {
    logWarning("updateRowStatus: " + error.message);
  }
}

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

function ensureStatusColumn(sheet) {
  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let statusColIndex = headers.indexOf("Statut activation");
    
    if (statusColIndex === -1) {
      const lastCol = sheet.getLastColumn();
      
      sheet.getRange(1, lastCol + 1).setValue("Statut activation");
      sheet.getRange(1, lastCol + 2).setValue("Date activation");
      sheet.getRange(1, lastCol + 3).setValue("Lien personnel");
      sheet.getRange(1, lastCol + 4).setValue("User ID");
      
      const headerRange = sheet.getRange(1, lastCol + 1, 1, 4);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#4f46e5");
      headerRange.setFontColor("#ffffff");
      
      statusColIndex = lastCol;
      
      logInfo("📋 Colonnes de suivi créées automatiquement");
    }
    
    return statusColIndex + 1;
    
  } catch (error) {
    logError("ensureStatusColumn", error);
    return sheet.getLastColumn() + 1;
  }
}

/**
 * ============================================
 * FONCTIONS DE MAINTENANCE
 * ============================================
 */

function manualProcessRow(rowNumber, sheetName) {
  try {
    logInfo(`🔧 Traitement manuel ligne ${rowNumber}...`);
    
    // Utilisation Sheet ID direct au lieu du parcours
    const sheetId = getSourceSheetId();
    
    if (!sheetId) {
      throw new Error("Sheet ID non configuré. Utilisez setSourceSheetId()");
    }
    
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    const sheet = sheetName 
      ? spreadsheet.getSheetByName(sheetName)
      : spreadsheet.getActiveSheet();
    
    if (!sheet) {
      throw new Error(`Feuille "${sheetName}" introuvable`);
    }
    
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

function reprocessUnfinishedRows(sheetName) {
  try {
    logWarning("🔄 Retraitement des lignes non finalisées...");
    
    const sheetId = getSourceSheetId();
    
    if (!sheetId) {
      throw new Error("Sheet ID non configuré");
    }
    
    const spreadsheet = SpreadsheetApp.openById(sheetId);
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
    
    for (let row = 2; row <= lastRow; row++) {
      try {
        if (isRowAlreadyProcessed(sheet, row)) {
          skipped++;
          continue;
        }
        
        processNewSubmission(sheet, row);
        processed++;
        
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

function getSystemReport() {
  try {
    const sheetId = getSourceSheetId();
    
    if (!sheetId) {
      return {
        error: "Sheet ID non configuré",
        status: "not_configured"
      };
    }
    
    const spreadsheet = SpreadsheetApp.openById(sheetId);
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

function analyzeSheet(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    const totalRows = lastRow - 1;
    
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
 * CONFIGURATION INITIALE
 * ============================================
 */

function setupNovaMail(deploymentId) {
  try {
    logInfo("🚀 Configuration NovaMail...");
    
    if (deploymentId) {
      setDeploymentId(deploymentId);
      logInfo("✅ Deployment ID configuré");
    }
    
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
    console.log("✅ Sheet ID: " + (getSourceSheetId() || "Non configuré"));
    console.log("✅ Sender Email: " + getDefaultSenderEmail());
    console.log("");
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
  
  // Test 2 : Sheet ID configuré
  const sheetId = getSourceSheetId();
  results.checks.push({ 
    name: "Sheet ID", 
    passed: !!sheetId,
    value: sheetId || "Non configuré"
  });
  
  // Test 3 : Deployment ID
  const deploymentId = getDeploymentId();
  results.checks.push({ 
    name: "Deployment ID", 
    passed: !!deploymentId,
    value: deploymentId || "Non configuré"
  });
  
  // Test 4 : Email expéditeur
  try {
    const sender = getDefaultSenderEmail();
    results.checks.push({
      name: "Email expéditeur",
      passed: !!sender && isValidEmail(sender),
      value: sender
    });
  } catch (e) {
    results.checks.push({
      name: "Email expéditeur",
      passed: false,
      error: e.message
    });
  }
  
  return results;
}
