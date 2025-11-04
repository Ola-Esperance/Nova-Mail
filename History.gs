/*****************************************************
 * NOVAMAIL SAAS - HISTORY.GS
 * ====================================================
 * Gestion centralisée de l'historique des campagnes
 * Stockage dans Google Sheets unique et réutilisable
 * 
 * @author NovaMail Team
 * @version 2.0.0
 * @lastModified 2025-11-03
 *****************************************************/

/**
 * ============================================
 * GESTION DU FICHIER D'HISTORIQUE
 * ============================================
 */

/**
 * 🗂️ Récupère ou crée le fichier Google Sheets d'historique unique
 * 
 * Logique :
 * 1. Vérifie si un ID est enregistré dans Script Properties
 * 2. Si oui, tente de l'ouvrir
 * 3. Si échec, recherche un fichier du même nom
 * 4. Si aucun trouvé, crée un nouveau fichier
 * 
 * @returns {Spreadsheet} Objet Spreadsheet Google Apps Script
 */
function getOrCreateHistoryFile() {
  try {
    // 1️⃣ Tentative de récupération depuis Script Properties
    let fileId = PropertiesService.getScriptProperties()
      .getProperty(HISTORY_CONFIG.PROPERTY_KEY);
    
    if (fileId) {
      try {
        const spreadsheet = SpreadsheetApp.openById(fileId);
        
        // Vérification que le fichier existe vraiment
        if (spreadsheet && spreadsheet.getId()) {
          logInfo(`📘 Fichier d'historique trouvé : ${fileId}`);
          return spreadsheet;
        }
      } catch (e) {
        // Fichier supprimé ou inaccessible
        logWarning(`⚠️ Fichier d'historique ${fileId} inaccessible - recréation...`);
        PropertiesService.getScriptProperties()
          .deleteProperty(HISTORY_CONFIG.PROPERTY_KEY);
      }
    }
    
    // 2️⃣ Recherche d'un fichier existant du même nom
    const existingFiles = DriveApp.getFilesByName(HISTORY_CONFIG.FILE_NAME);
    
    if (existingFiles.hasNext()) {
      const file = existingFiles.next();
      const spreadsheet = SpreadsheetApp.openById(file.getId());
      
      // Enregistrement de l'ID pour usage futur
      PropertiesService.getScriptProperties()
        .setProperty(HISTORY_CONFIG.PROPERTY_KEY, file.getId());
      
      logInfo(`📄 Fichier d'historique existant réutilisé : ${file.getId()}`);
      return spreadsheet;
    }
    
    // 3️⃣ Création d'un nouveau fichier
    const newSpreadsheet = SpreadsheetApp.create(HISTORY_CONFIG.FILE_NAME);
    
    // Enregistrement de l'ID
    PropertiesService.getScriptProperties()
      .setProperty(HISTORY_CONFIG.PROPERTY_KEY, newSpreadsheet.getId());
    
    logInfo(`🆕 Nouveau fichier d'historique créé : ${newSpreadsheet.getId()}`);
    
    return newSpreadsheet;
    
  } catch (error) {
    logError("getOrCreateHistoryFile", error);
    throw new Error("Impossible de créer/accéder au fichier d'historique : " + error.message);
  }
}

/**
 * 📄 Récupère ou crée la feuille "Historique_Campagnes"
 * 
 * @returns {Sheet} Objet Sheet Google Apps Script
 */
function getHistorySheet() {
  try {
    const spreadsheet = getOrCreateHistoryFile();
    
    // Tentative de récupération de la feuille existante
    let sheet = spreadsheet.getSheetByName(HISTORY_CONFIG.SHEET_NAME);
    
    if (!sheet) {
      // Création de la feuille avec en-têtes
      sheet = spreadsheet.insertSheet(HISTORY_CONFIG.SHEET_NAME);
      
      // Ajout des en-têtes
      const headers = [
        "Date d'envoi",
        "Type",
        "Nom de la campagne",
        "Objet",
        "Nombre de destinataires",
        "Emails (aperçu)",
        "Statut",
        "Détails / Erreur"
      ];
      
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Formatage des en-têtes
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#4f46e5");
      headerRange.setFontColor("#ffffff");
      
      // Ajustement largeur colonnes
      sheet.setColumnWidth(1, 150); // Date
      sheet.setColumnWidth(2, 100); // Type
      sheet.setColumnWidth(3, 200); // Nom
      sheet.setColumnWidth(4, 250); // Objet
      sheet.setColumnWidth(5, 80);  // Nb destinataires
      sheet.setColumnWidth(6, 300); // Emails
      sheet.setColumnWidth(7, 100); // Statut
      sheet.setColumnWidth(8, 300); // Détails
      
      // Freeze première ligne
      sheet.setFrozenRows(1);
      
      logInfo("📋 Nouvelle feuille d'historique créée avec en-têtes");
    }
    
    return sheet;
    
  } catch (error) {
    logError("getHistorySheet", error);
    throw error;
  }
}

/**
 * ============================================
 * ENREGISTREMENT DANS L'HISTORIQUE
 * ============================================
 */

/**
 * 📝 Enregistre une campagne dans l'historique
 * 
 * Cette fonction est appelée après chaque envoi (direct, planifié, test)
 * 
 * @param {string} campaignName - Nom de la campagne
 * @param {Array<Object>|Array<string>} recipients - Destinataires
 * @param {Date|string} sendDate - Date d'envoi
 * @param {string} status - Statut (Envoyé, Partiel, Erreur, Test)
 * @param {string} details - Détails supplémentaires
 * @param {string} subject - Objet de l'email
 * @param {string} type - Type de campagne (direct, scheduled, test)
 */
function logCampaignHistory(campaignName, recipients, sendDate, status, details, subject, type) {
  try {
    const sheet = getHistorySheet();
    
    // ✅ Sécurisation de la date
    let dateToLog;
    if (sendDate instanceof Date) {
      dateToLog = sendDate;
    } else if (typeof sendDate === "string") {
      dateToLog = new Date(sendDate);
    } else {
      dateToLog = new Date();
    }
    
    // Validation de la date
    if (isNaN(dateToLog.getTime())) {
      dateToLog = new Date();
      details += " (⚠️ Date invalide corrigée)";
    }
    
    // ✅ Formatage de la date
    const formattedDate = Utilities.formatDate(
      dateToLog,
      Session.getScriptTimeZone(),
      "dd/MM/yyyy HH:mm:ss"
    );
    
    // ✅ Gestion des destinataires
    let recipientCount = 0;
    let emailsPreview = "";
    
    if (Array.isArray(recipients)) {
      recipientCount = recipients.length;
      
      // Aperçu des emails (max 3)
      const emailsList = recipients
        .slice(0, 3)
        .map(r => {
          if (typeof r === "string") return r;
          if (r.email) return r.email;
          return "email_inconnu";
        });
      
      emailsPreview = emailsList.join(", ");
      
      if (recipients.length > 3) {
        emailsPreview += ` (+ ${recipients.length - 3} autres)`;
      }
    } else if (typeof recipients === "string") {
      recipientCount = 1;
      emailsPreview = recipients;
    }
    
    // ✅ Détermination du type
    const campaignType = type || (campaignName.includes("Test") ? "Test" : "Directe");
    
    // ✅ Construction de la ligne
    const row = [
      formattedDate,
      campaignType,
      campaignName || "(Sans nom)",
      subject || "(Sans objet)",
      recipientCount,
      emailsPreview,
      status || "Inconnu",
      details || ""
    ];
    
    // ✅ Ajout de la ligne
    sheet.appendRow(row);
    
    // ✅ Formatage conditionnel selon statut
    const lastRow = sheet.getLastRow();
    const statusCell = sheet.getRange(lastRow, 7); // Colonne "Statut"
    
    switch (status) {
      case "Envoyé":
        statusCell.setBackground("#d1fae5"); // Vert clair
        statusCell.setFontColor("#065f46");
        break;
      case "Partiel":
        statusCell.setBackground("#fef3c7"); // Jaune clair
        statusCell.setFontColor("#92400e");
        break;
      case "Erreur":
        statusCell.setBackground("#fee2e2"); // Rouge clair
        statusCell.setFontColor("#991b1b");
        break;
      case "Test":
        statusCell.setBackground("#e0e7ff"); // Bleu clair
        statusCell.setFontColor("#3730a3");
        break;
    }
    
    logInfo(
      `📊 Historique : ${campaignName} (${status}) - ` +
      `${recipientCount} destinataire(s)`
    );
    
  } catch (error) {
    // Log silencieux : l'historique est optionnel
    logError("logCampaignHistory", error);
  }
}

/**
 * ============================================
 * RÉCUPÉRATION DE L'HISTORIQUE
 * ============================================
 */

/**
 * 📤 Récupère l'historique complet pour affichage frontend
 * 
 * @param {number} limit - Nombre max de lignes (0 = toutes)
 * @returns {Array<Object>} Historique formaté
 * 
 * @example
 * const history = getHistoryData(50); // 50 dernières campagnes
 */
function getHistoryData(limit) {
  try {
    const sheet = getHistorySheet();
    const data = sheet.getDataRange().getValues();
    
    if (!data || data.length <= 1) {
      // Seulement en-têtes ou vide
      return [];
    }
    
    const headers = data[0];
    const rows = data.slice(1); // Ignorer les en-têtes
    
    // Mapping des colonnes
    const mapping = {
      "Date d'envoi": "date",
      "Type": "type",
      "Nom de la campagne": "name",
      "Objet": "subject",
      "Nombre de destinataires": "recipients",
      "Emails (aperçu)": "emailsPreview",
      "Statut": "status",
      "Détails / Erreur": "details"
    };
    
    // Conversion en objets
    const history = rows.map(row => {
      const entry = {};
      headers.forEach((header, index) => {
        const key = mapping[header] || header;
        entry[key] = row[index];
      });
      return entry;
    });
    
    // Tri par date décroissante (plus récent en premier)
    history.sort((a, b) => {
      const dateA = new Date(a.date);
      const dateB = new Date(b.date);
      return dateB - dateA;
    });
    
    // Limitation si demandée
    if (limit && limit > 0) {
      return history.slice(0, limit);
    }
    
    return history;
    
  } catch (error) {
    logError("getHistoryData", error);
    return [];
  }
}

/**
 * 📊 Récupère des statistiques sur l'historique
 * 
 * @returns {Object} Statistiques {total, sent, failed, lastWeek}
 */
function getHistoryStats() {
  try {
    const history = getHistoryData(0); // Tout l'historique
    
    const stats = {
      totalCampaigns: history.length,
      totalRecipients: 0,
      successfulCampaigns: 0,
      failedCampaigns: 0,
      partialCampaigns: 0,
      testCampaigns: 0,
      lastWeekCampaigns: 0,
      lastMonthCampaigns: 0
    };
    
    const now = new Date();
    const oneWeekAgo = new Date(now - 7 * 24 * 60 * 60 * 1000);
    const oneMonthAgo = new Date(now - 30 * 24 * 60 * 60 * 1000);
    
    history.forEach(entry => {
      // Comptage destinataires
      if (entry.recipients && !isNaN(entry.recipients)) {
        stats.totalRecipients += parseInt(entry.recipients, 10);
      }
      
      // Comptage par statut
      switch (entry.status) {
        case "Envoyé":
          stats.successfulCampaigns++;
          break;
        case "Erreur":
          stats.failedCampaigns++;
          break;
        case "Partiel":
          stats.partialCampaigns++;
          break;
        case "Test":
          stats.testCampaigns++;
          break;
      }
      
      // Comptage temporel
      try {
        const entryDate = new Date(entry.date);
        if (entryDate >= oneWeekAgo) {
          stats.lastWeekCampaigns++;
        }
        if (entryDate >= oneMonthAgo) {
          stats.lastMonthCampaigns++;
        }
      } catch (e) {
        // Date invalide, ignorer
      }
    });
    
    return stats;
    
  } catch (error) {
    logError("getHistoryStats", error);
    return {
      totalCampaigns: 0,
      totalRecipients: 0,
      successfulCampaigns: 0,
      failedCampaigns: 0
    };
  }
}

/**
 * ============================================
 * NETTOYAGE DE L'HISTORIQUE
 * ============================================
 */

/**
 * 🧹 Efface tout l'historique (garde les en-têtes)
 * 
 * @returns {boolean} Succès
 */
function clearHistorySheet() {
  try {
    const sheet = getHistorySheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow > 1) {
      // Suppression de toutes les lignes sauf les en-têtes
      sheet.deleteRows(2, lastRow - 1);
      logInfo("🧹 Historique vidé (en-têtes conservés)");
    } else {
      logInfo("📋 Historique déjà vide");
    }
    
    return true;
    
  } catch (error) {
    logError("clearHistorySheet", error);
    return false;
  }
}

/**
 * 🗑️ Supprime les entrées plus anciennes qu'une date donnée
 * 
 * @param {Date} cutoffDate - Date limite (plus ancien = supprimé)
 * @returns {number} Nombre de lignes supprimées
 */
function purgeOldHistory(cutoffDate) {
  try {
    const sheet = getHistorySheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return 0; // Seulement en-têtes
    
    let deletedCount = 0;
    
    // Parcourir de bas en haut pour ne pas décaler les indices
    for (let i = data.length - 1; i >= 1; i--) {
      try {
        const dateStr = data[i][0]; // Première colonne = date
        const entryDate = new Date(dateStr);
        
        if (!isNaN(entryDate.getTime()) && entryDate < cutoffDate) {
          sheet.deleteRow(i + 1); // +1 car indices Google Sheets commencent à 1
          deletedCount++;
        }
      } catch (e) {
        // Date invalide, ignorer
      }
    }
    
    logInfo(`🗑️ ${deletedCount} entrée(s) ancienne(s) supprimée(s)`);
    
    return deletedCount;
    
  } catch (error) {
    logError("purgeOldHistory", error);
    return 0;
  }
}

/**
 * 🧹 Nettoie les entrées de test uniquement
 * 
 * @returns {number} Nombre de tests supprimés
 */
function clearTestHistory() {
  try {
    const sheet = getHistorySheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return 0;
    
    let deletedCount = 0;
    
    // Parcourir de bas en haut
    for (let i = data.length - 1; i >= 1; i--) {
      const type = data[i][1]; // Colonne "Type"
      const status = data[i][6]; // Colonne "Statut"
      
      if (type === "Test" || status === "Test") {
        sheet.deleteRow(i + 1);
        deletedCount++;
      }
    }
    
    logInfo(`🧹 ${deletedCount} test(s) supprimé(s) de l'historique`);
    
    return deletedCount;
    
  } catch (error) {
    logError("clearTestHistory", error);
    return 0;
  }
}

/**
 * ============================================
 * EXPORT DE L'HISTORIQUE
 * ============================================
 */

/**
 * 📥 Exporte l'historique au format CSV
 * 
 * @returns {string} Contenu CSV
 */
function exportHistoryToCSV() {
  try {
    const history = getHistoryData(0); // Tout
    
    if (!history || history.length === 0) {
      return "Date,Type,Campagne,Objet,Destinataires,Statut,Détails\n";
    }
    
    // En-têtes
    let csv = "Date,Type,Campagne,Objet,Destinataires,Statut,Détails\n";
    
    // Lignes
    history.forEach(entry => {
      csv += [
        `"${entry.date || ""}"`,
        `"${entry.type || ""}"`,
        `"${entry.name || ""}"`,
        `"${entry.subject || ""}"`,
        entry.recipients || 0,
        `"${entry.status || ""}"`,
        `"${(entry.details || "").replace(/"/g, '""')}"` // Échappement guillemets
      ].join(",") + "\n";
    });
    
    return csv;
    
  } catch (error) {
    logError("exportHistoryToCSV", error);
    return "";
  }
}

/**
 * ============================================
 * FONCTIONS DEBUG
 * ============================================
 */

/**
 * 🔍 Affiche l'ID du fichier d'historique actuel
 * 
 * @returns {string} ID du fichier ou message
 */
function debugShowHistoryFileId() {
  const id = PropertiesService.getScriptProperties()
    .getProperty(HISTORY_CONFIG.PROPERTY_KEY);
  
  if (id) {
    Logger.log(`📘 ID fichier d'historique : ${id}`);
    Logger.log(`🔗 URL : https://docs.google.com/spreadsheets/d/${id}/edit`);
    return id;
  } else {
    Logger.log("⚠️ Aucun fichier d'historique enregistré");
    return null;
  }
}

/**
 * 🧼 Réinitialise manuellement l'ID d'historique (debug)
 * ⚠️ Le prochain appel recréera ou trouvera un fichier
 */
function resetHistoryFileId() {
  PropertiesService.getScriptProperties()
    .deleteProperty(HISTORY_CONFIG.PROPERTY_KEY);
  
  logWarning("🧽 ID d'historique supprimé → recréation au prochain appel");
}
