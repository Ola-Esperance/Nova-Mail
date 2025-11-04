/*****************************************************
 * NOVAMAIL SAAS - SCHEDULING.GS
 * ====================================================
 * Gestion des campagnes planifiées et récurrentes
 * Déclencheur maître pour exécution automatique
 * 
 * @author NovaMail Team
 * @version 2.0.0
 * @lastModified 2025-11-03
 *****************************************************/

/**
 * ============================================
 * CONSTANTES PLANIFICATION
 * ============================================
 */

const SCHEDULING_CONFIG = {
  // Nom du trigger maître
  MASTER_TRIGGER_FUNCTION: "runAllScheduledCampaigns",
  
  // Intervalle de vérification (minutes)
  CHECK_INTERVAL_MINUTES: 1,
  
  // Marge de tolérance pour exécution (secondes)
  EXECUTION_TOLERANCE_SECONDS: 120
};

/**
 * ============================================
 * PLANIFICATION DE CAMPAGNES
 * ============================================
 */

/**
 * 📅 Planifie l'envoi d'une campagne à une date future
 * 
 * Vérifie les quotas mais ne les consomme PAS (consommation à l'envoi)
 * 
 * @param {Object} campaignInput - Données campagne
 * @param {Array<Object>} campaignInput.recipients - Destinataires
 * @param {string} campaignInput.subject - Objet
 * @param {string} campaignInput.htmlBody - Contenu HTML
 * @param {string} campaignInput.sendAt - Date ISO planifiée
 * @param {Array<Object>} campaignInput.attachments - Pièces jointes (optionnel)
 * @returns {Object} Résultat {success, id, scheduledAt}
 * 
 * @example
 * scheduleCampaign({
 *   recipients: [{nom: "John", email: "john@example.com"}],
 *   subject: "Newsletter",
 *   htmlBody: "<p>Bonjour {{nom}}</p>",
 *   sendAt: "2025-12-25T10:00:00Z",
 *   attachments: []
 * });
 */
function scheduleCampaign(campaignInput) {
  try {
    // ✅ Validation des paramètres
    validateScheduleInput(campaignInput);
    
    // ✅ Vérification de la date
    const sendDate = new Date(campaignInput.sendAt);
    
    if (isNaN(sendDate.getTime())) {
      throw new Error("Date de planification invalide");
    }
    
    if (sendDate <= new Date()) {
      throw new Error("La date de planification doit être dans le futur");
    }
    
    // ✅ Vérification des permissions selon version
    const config = getVersionConfig();
    
    if (!config.scheduleSend) {
      throw new Error(
        `❌ Planification non disponible avec votre version ${config.name}. ` +
        `Passez à STARTER ou supérieur.`
      );
    }
    
    // Vérification limite 48h pour STARTER
    if (config.scheduleSend === "limited") {
      const maxDate = new Date(Date.now() + (48 * 60 * 60 * 1000)); // +48h
      if (sendDate > maxDate) {
        throw new Error(
          `❌ Version ${config.name} : planification limitée à 48h maximum. ` +
          `Passez à PRO pour planifier sans limite.`
        );
      }
    }
    
    // ✅ Vérification des quotas (sans consommation)
    checkQuotaAvailable(campaignInput.recipients.length);
    
    // ✅ Récupération informations client
    const client = loadClientByCurrentUser();
    
    // ✅ Construction de la campagne planifiée
    const scheduledCampaign = {
      id: STORAGE_PREFIX.SCHEDULED_CAMPAIGN + generateFullId(),
      type: "scheduled",
      name: campaignInput.subject || "Campagne planifiée",
      subject: campaignInput.subject,
      htmlBody: campaignInput.htmlBody,
      recipients: campaignInput.recipients,
      attachments: campaignInput.attachments || [],
      sendAt: sendDate.toISOString(),
      createdAt: new Date().toISOString(),
      status: "pending",
      
      // Informations expéditeur
      senderEmail: client ? client.senderEmail : DEFAULT_SENDER_EMAIL,
      senderName: client ? client.companyName || client.fullName : DEFAULT_SENDER_NAME,
      replyTo: client ? client.replyEmail : DEFAULT_SENDER_EMAIL,
      
      // Métadonnées
      userId: client ? client.userId : null,
      createdBy: Session.getActiveUser().getEmail()
    };
    
    // ✅ Sauvegarde dans Script Properties
    saveScriptProperty(scheduledCampaign.id, scheduledCampaign);
    
    // ✅ Activation du trigger maître (si pas déjà actif)
    ensureMasterTriggerActive();
    
    logInfo(
      `📅 Campagne planifiée : "${scheduledCampaign.name}" ` +
      `pour ${formatDateFR(sendDate)} (ID: ${scheduledCampaign.id})`
    );
    
    return {
      success: true,
      id: scheduledCampaign.id,
      scheduledAt: formatDateFR(sendDate),
      message: `Campagne planifiée avec succès pour le ${formatDateFR(sendDate)}`
    };
    
  } catch (error) {
    logError("scheduleCampaign", error);
    throw error;
  }
}

/**
 * Valide les données d'entrée pour planification
 * 
 * @param {Object} input - Données à valider
 * @throws {Error} Si validation échoue
 */
function validateScheduleInput(input) {
  if (!input || typeof input !== "object") {
    throw new Error("Données de campagne invalides");
  }
  
  const required = ["recipients", "subject", "htmlBody", "sendAt"];
  const missing = required.filter(key => !input[key]);
  
  if (missing.length > 0) {
    throw new Error(`Champs requis manquants : ${missing.join(", ")}`);
  }
  
  if (!Array.isArray(input.recipients) || input.recipients.length === 0) {
    throw new Error("Au moins un destinataire est requis");
  }
  
  return true;
}

/**
 * ============================================
 * EXÉCUTION DES CAMPAGNES PLANIFIÉES
 * ============================================
 */

/**
 * ⏰ Trigger maître : vérifie et exécute les campagnes arrivées à échéance
 * 
 * Cette fonction est appelée automatiquement toutes les minutes
 * Elle est IDEMPOTENTE : peut être rappelée sans risque
 */
function runAllScheduledCampaigns() {
  try {
    const now = new Date();
    logInfo(`⏰ Vérification campagnes planifiées (${formatDateShort(now)})`);
    
    // Récupération de toutes les campagnes planifiées
    const allProps = PropertiesService.getScriptProperties().getProperties();
    const scheduledCampaigns = [];
    
    for (const key in allProps) {
      if (key.startsWith(STORAGE_PREFIX.SCRIPT_PROPS + ":" + STORAGE_PREFIX.SCHEDULED_CAMPAIGN)) {
        try {
          const campaign = JSON.parse(allProps[key]);
          if (campaign.status === "pending") {
            scheduledCampaigns.push(campaign);
          }
        } catch (e) {
          logWarning(`Campagne corrompue ignorée : ${key}`);
        }
      }
    }
    
    if (scheduledCampaigns.length === 0) {
      logInfo("💤 Aucune campagne planifiée à exécuter");
      return;
    }
    
    logInfo(`🔍 ${scheduledCampaigns.length} campagne(s) planifiée(s) en attente`);
    
    // Tri par date croissante
    scheduledCampaigns.sort((a, b) => 
      new Date(a.sendAt) - new Date(b.sendAt)
    );
    
    // Vérification et exécution
    let executedCount = 0;
    
    scheduledCampaigns.forEach(campaign => {
      const sendDate = new Date(campaign.sendAt);
      const timeDiff = (now - sendDate) / 1000; // Différence en secondes
      
      // Vérifier si la date est atteinte (avec tolérance)
      if (timeDiff >= -SCHEDULING_CONFIG.EXECUTION_TOLERANCE_SECONDS) {
        try {
          executeScheduledCampaign(campaign);
          executedCount++;
        } catch (err) {
          logError(`Erreur exécution campagne ${campaign.id}`, err);
          markCampaignAsFailed(campaign.id, err.message);
        }
      }
    });
    
    if (executedCount > 0) {
      logInfo(`✅ ${executedCount} campagne(s) exécutée(s)`);
    } else {
      logInfo("💤 Aucune campagne à exécuter pour le moment");
    }
    
  } catch (error) {
    logError("runAllScheduledCampaigns", error);
  }
}

/**
 * 🚀 Exécute une campagne planifiée spécifique
 * 
 * @param {Object} campaign - Objet campagne planifiée
 */
function executeScheduledCampaign(campaign) {
  try {
    logInfo(`🚀 Exécution campagne planifiée : ${campaign.name}`);
    
    // ✅ Vérification des quotas au moment de l'envoi
    checkQuotaAvailable(campaign.recipients.length);
    
    // ✅ Conversion des pièces jointes
    const blobs = [];
    
    if (campaign.attachments && campaign.attachments.length > 0) {
      campaign.attachments.forEach(att => {
        try {
          blobs.push(base64ToBlob(att.base64, att.name, att.mime));
        } catch (err) {
          logWarning(`Pièce jointe ignorée : ${att.name} - ${err.message}`);
        }
      });
    }
    
    // ✅ Envoi par batches
    let sentCount = 0;
    const errors = [];
    
    for (let i = 0; i < campaign.recipients.length; i += GMAIL_BATCH_SIZE) {
      const batch = campaign.recipients.slice(i, i + GMAIL_BATCH_SIZE);
      
      batch.forEach(recipient => {
        try {
          const personalizedSubject = replacePlaceholders(campaign.subject, recipient);
          const personalizedBody = replacePlaceholders(campaign.htmlBody, recipient);
          
          const mailOptions = {
            htmlBody: personalizedBody,
            name: campaign.senderName || DEFAULT_SENDER_NAME,
            replyTo: campaign.replyTo || DEFAULT_SENDER_EMAIL
          };
          
          if (blobs.length > 0) {
            mailOptions.attachments = blobs;
          }
          
          GmailApp.sendEmail(
            recipient.email,
            personalizedSubject,
            stripHtml(personalizedBody),
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
      
      // Délai entre batches
      if (i + GMAIL_BATCH_SIZE < campaign.recipients.length) {
        Utilities.sleep(GMAIL_BATCH_DELAY_MS);
      }
    }
    
    // ✅ Incrémentation des quotas
    if (sentCount > 0) {
      incrementQuota(sentCount);
    }
    
    // ✅ Enregistrement dans l'historique
    logCampaignHistory(
      campaign.name,
      campaign.recipients,
      new Date(campaign.sendAt),
      errors.length === 0 ? "Envoyé" : "Partiel",
      errors.length > 0 ? `${errors.length} erreur(s)` : "Succès (planifié)",
      campaign.subject
    );
    
    // ✅ Suppression de la campagne planifiée
    deleteScriptProperty(campaign.id);
    
    logInfo(
      `✅ Campagne "${campaign.name}" exécutée : ` +
      `${sentCount}/${campaign.recipients.length} envoyés`
    );
    
  } catch (error) {
    logError("executeScheduledCampaign", error);
    throw error;
  }
}

/**
 * Marque une campagne comme échouée
 * 
 * @param {string} campaignId - ID de la campagne
 * @param {string} errorMessage - Message d'erreur
 */
function markCampaignAsFailed(campaignId, errorMessage) {
  try {
    const campaign = loadScriptProperty(campaignId);
    
    if (campaign) {
      campaign.status = "failed";
      campaign.error = errorMessage;
      campaign.failedAt = new Date().toISOString();
      
      saveScriptProperty(campaignId, campaign);
      
      // Log dans l'historique
      logCampaignHistory(
        campaign.name,
        campaign.recipients || [],
        new Date(),
        "Erreur",
        errorMessage,
        campaign.subject
      );
    }
    
  } catch (error) {
    logError("markCampaignAsFailed", error);
  }
}

/**
 * ============================================
 * GESTION DES CAMPAGNES PLANIFIÉES
 * ============================================
 */

/**
 * 📋 Liste toutes les campagnes planifiées (pending)
 * 
 * @returns {Array<Object>} Liste des campagnes triées par date
 */
function getScheduledCampaigns() {
  try {
    const allProps = PropertiesService.getScriptProperties().getProperties();
    const campaigns = [];
    const prefix = STORAGE_PREFIX.SCRIPT_PROPS + ":" + STORAGE_PREFIX.SCHEDULED_CAMPAIGN;
    
    for (const key in allProps) {
      if (key.startsWith(prefix)) {
        try {
          const campaign = JSON.parse(allProps[key]);
          if (campaign.status === "pending") {
            campaigns.push({
              id: campaign.id,
              name: campaign.name,
              subject: campaign.subject,
              sendAt: campaign.sendAt,
              scheduledAt: campaign.sendAt, // Alias pour compatibilité frontend
              recipients: campaign.recipients,
              createdAt: campaign.createdAt,
              recipientCount: campaign.recipients ? campaign.recipients.length : 0
            });
          }
        } catch (e) {
          logWarning(`Campagne corrompue ignorée : ${key}`);
        }
      }
    }
    
    // Tri par date d'envoi croissante
    campaigns.sort((a, b) => new Date(a.sendAt) - new Date(b.sendAt));
    
    logInfo(`📋 ${campaigns.length} campagne(s) planifiée(s) listée(s)`);
    
    return campaigns;
    
  } catch (error) {
    logError("getScheduledCampaigns", error);
    return [];
  }
}

/**
 * 🔄 Modifie la date de planification d'une campagne
 * 
 * @param {string} campaignId - ID de la campagne
 * @param {string} newSendAtIso - Nouvelle date ISO
 * @returns {Object} Résultat {success, scheduledAt}
 */
function updateScheduledDate(campaignId, newSendAtIso) {
  if (!campaignId || !newSendAtIso) {
    throw new Error("ID campagne et nouvelle date requis");
  }
  
  try {
    // Chargement de la campagne
    const campaign = loadScriptProperty(campaignId);
    
    if (!campaign) {
      throw new Error(`Campagne introuvable : ${campaignId}`);
    }
    
    if (campaign.status !== "pending") {
      throw new Error("Impossible de modifier une campagne déjà exécutée");
    }
    
    // Validation de la nouvelle date
    const newDate = new Date(newSendAtIso);
    
    if (isNaN(newDate.getTime())) {
      throw new Error("Date invalide : " + newSendAtIso);
    }
    
    if (newDate <= new Date()) {
      throw new Error("La nouvelle date doit être dans le futur");
    }
    
    // Mise à jour
    campaign.sendAt = newDate.toISOString();
    campaign.updatedAt = new Date().toISOString();
    
    saveScriptProperty(campaignId, campaign);
    
    logInfo(`🔄 Campagne replanifiée : ${campaignId} → ${formatDateFR(newDate)}`);
    
    return {
      success: true,
      id: campaignId,
      scheduledAt: formatDateFR(newDate),
      message: `Campagne replanifiée pour le ${formatDateFR(newDate)}`
    };
    
  } catch (error) {
    logError("updateScheduledDate", error);
    throw error;
  }
}

/**
 * 🗑️ Supprime une campagne planifiée
 * 
 * @param {string} campaignId - ID de la campagne
 * @returns {Object} Résultat {success, id}
 */
function deleteScheduledCampaign(campaignId) {
  if (!campaignId) {
    throw new Error("ID campagne requis");
  }
  
  try {
    const campaign = loadScriptProperty(campaignId);
    
    if (!campaign) {
      throw new Error(`Campagne introuvable : ${campaignId}`);
    }
    
    // Suppression
    deleteScriptProperty(campaignId);
    
    logInfo(`🗑️ Campagne planifiée supprimée : ${campaign.name} (${campaignId})`);
    
    return {
      success: true,
      id: campaignId,
      message: "Campagne supprimée avec succès"
    };
    
  } catch (error) {
    logError("deleteScheduledCampaign", error);
    throw error;
  }
}

/**
 * ============================================
 * GESTION DU TRIGGER MAÎTRE
 * ============================================
 */

/**
 * ✅ Active le trigger maître s'il n'existe pas déjà
 * 
 * @returns {boolean} true si trigger activé
 */
function ensureMasterTriggerActive() {
  try {
    // Vérification si trigger existe déjà
    const triggers = ScriptApp.getProjectTriggers();
    const existingTrigger = triggers.find(t => 
      t.getHandlerFunction() === SCHEDULING_CONFIG.MASTER_TRIGGER_FUNCTION
    );
    
    if (existingTrigger) {
      logInfo("✅ Trigger maître déjà actif");
      return true;
    }
    
    // Création du trigger
    ScriptApp.newTrigger(SCHEDULING_CONFIG.MASTER_TRIGGER_FUNCTION)
      .timeBased()
      .everyMinutes(SCHEDULING_CONFIG.CHECK_INTERVAL_MINUTES)
      .create();
    
    logInfo(
      `✅ Trigger maître créé : vérification toutes les ` +
      `${SCHEDULING_CONFIG.CHECK_INTERVAL_MINUTES} minute(s)`
    );
    
    return true;
    
  } catch (error) {
    logError("ensureMasterTriggerActive", error);
    return false;
  }
}

/**
 * 🗑️ Supprime le trigger maître
 * Utile pour maintenance ou désactivation temporaire
 * 
 * @returns {boolean} Succès
 */
function removeMasterTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let removed = false;
    
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === SCHEDULING_CONFIG.MASTER_TRIGGER_FUNCTION) {
        ScriptApp.deleteTrigger(trigger);
        removed = true;
      }
    });
    
    if (removed) {
      logInfo("🗑️ Trigger maître supprimé");
    } else {
      logWarning("⚠️ Aucun trigger maître trouvé");
    }
    
    return removed;
    
  } catch (error) {
    logError("removeMasterTrigger", error);
    return false;
  }
}

/**
 * 🔄 Réinstalle le trigger maître (supprime puis recrée)
 * 
 * @returns {boolean} Succès
 */
function reinstallMasterTrigger() {
  try {
    removeMasterTrigger();
    Utilities.sleep(1000); // Délai de sécurité
    return ensureMasterTriggerActive();
  } catch (error) {
    logError("reinstallMasterTrigger", error);
    return false;
  }
}

/**
 * ============================================
 * FONCTIONS DE DEBUG
 * ============================================
 */

/**
 * 🔧 Force l'exécution immédiate d'une campagne planifiée (debug)
 * 
 * @param {string} campaignId - ID de la campagne
 * @returns {Object} Résultat de l'exécution
 */
function forceExecuteCampaign(campaignId) {
  try {
    const campaign = loadScriptProperty(campaignId);
    
    if (!campaign) {
      throw new Error(`Campagne introuvable : ${campaignId}`);
    }
    
    logWarning(`⚠️ Exécution forcée de : ${campaign.name}`);
    
    executeScheduledCampaign(campaign);
    
    return {
      success: true,
      message: `Campagne "${campaign.name}" exécutée avec succès`
    };
    
  } catch (error) {
    logError("forceExecuteCampaign", error);
    throw error;
  }
}

/**
 * 🧪 Test du système de planification
 * 
 * @returns {Object} Résultat du test
 */
function testSchedulingSystem() {
  const results = {
    success: true,
    tests: []
  };
  
  // Test 1 : Trigger existe
  const triggers = ScriptApp.getProjectTriggers();
  const hasTrigger = triggers.some(t => 
    t.getHandlerFunction() === SCHEDULING_CONFIG.MASTER_TRIGGER_FUNCTION
  );
  
  results.tests.push({
    name: "Trigger maître",
    passed: hasTrigger,
    message: hasTrigger ? "✅ Actif" : "❌ Absent"
  });
  
  // Test 2 : Campagnes planifiées
  const scheduled = getScheduledCampaigns();
  results.tests.push({
    name: "Campagnes planifiées",
    passed: true,
    message: `${scheduled.length} campagne(s) en attente`
  });
  
  // Test 3 : Permissions
  try {
    GmailApp.getAliases();
    results.tests.push({
      name: "Permissions Gmail",
      passed: true,
      message: "✅ OK"
    });
  } catch (e) {
    results.tests.push({
      name: "Permissions Gmail",
      passed: false,
      message: "❌ Manquantes"
    });
    results.success = false;
  }
  
  return results;
}
