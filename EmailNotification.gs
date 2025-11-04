/*****************************************************
 * NOVAMAIL SAAS - EMAILNOTIFICATION.GS
 * ====================================================
 * ✅ FIX : Système de notification email robuste
 * ✅ Logs détaillés à chaque étape
 * ✅ Fallback automatique développeur si erreur
 * 
 * @version 1.0.0
 * @lastModified 2025-11-04
 *****************************************************/

/**
 * ============================================
 * ENVOI EMAIL AVEC LOGS DÉTAILLÉS
 * ============================================
 */

/**
 * 📧 Envoie un email avec tracking complet et gestion d'erreurs
 * 
 * @param {Object} params - Paramètres d'envoi
 * @param {string} params.to - Email destinataire
 * @param {string} params.subject - Sujet
 * @param {string} params.htmlBody - Contenu HTML
 * @param {string} params.textBody - Contenu texte (optionnel)
 * @param {string} params.from - Email expéditeur (optionnel)
 * @param {string} params.fromName - Nom expéditeur (optionnel)
 * @param {string} params.replyTo - Email de réponse (optionnel)
 * @param {Array<Blob>} params.attachments - Pièces jointes (optionnel)
 * @returns {Object} Résultat {success, message, timestamp}
 */
function sendEmailWithTracking(params) {
  const startTime = new Date();
  const logId = generateShortId();
  
  try {
    // ===== ÉTAPE 1 : VALIDATION =====
    logInfo(`[${logId}] 📧 Démarrage envoi email...`);
    
    if (!params || !params.to) {
      throw new Error("Paramètre 'to' (destinataire) manquant");
    }
    
    if (!isValidEmail(params.to)) {
      throw new Error(`Email destinataire invalide : ${params.to}`);
    }
    
    if (!params.subject || !params.htmlBody) {
      throw new Error("Sujet ou contenu manquant");
    }
    
    logInfo(`[${logId}] ✅ Validation OK - Destinataire: ${params.to}`);
    
    // ===== ÉTAPE 2 : PRÉPARATION =====
    const senderEmail = params.from || getDefaultSenderEmail();
    const senderName = params.fromName || DEFAULT_SENDER_NAME;
    const replyTo = params.replyTo || senderEmail;
    
    // Version texte automatique si absente
    const textBody = params.textBody || stripHtml(params.htmlBody);
    
    logInfo(`[${logId}] 📤 Expéditeur: ${senderName} <${senderEmail}>`);
    logInfo(`[${logId}] 📬 Réponse à: ${replyTo}`);
    logInfo(`[${logId}] 📎 Pièces jointes: ${params.attachments ? params.attachments.length : 0}`);
    
    // ===== ÉTAPE 3 : OPTIONS GMAIL =====
    const mailOptions = {
      htmlBody: params.htmlBody,
      name: senderName,
      replyTo: replyTo
    };
    
    // Ajout pièces jointes si présentes
    if (params.attachments && params.attachments.length > 0) {
      mailOptions.attachments = params.attachments;
      logInfo(`[${logId}] 📎 ${params.attachments.length} pièce(s) jointe(s) ajoutée(s)`);
    }
    
    // ===== ÉTAPE 4 : ENVOI =====
    logInfo(`[${logId}] 🚀 Envoi en cours via GmailApp...`);
    
    GmailApp.sendEmail(
      params.to,
      params.subject,
      textBody,
      mailOptions
    );
    
    // ===== ÉTAPE 5 : CONFIRMATION =====
    const duration = new Date() - startTime;
    const result = {
      success: true,
      message: `Email envoyé avec succès à ${params.to}`,
      timestamp: new Date().toISOString(),
      duration: duration,
      logId: logId
    };
    
    logInfo(`[${logId}] ✅ Envoi réussi en ${duration}ms`);
    
    // Enregistrement dans historique (silencieux)
    try {
      logEmailSent(params.to, params.subject, true, duration);
    } catch (e) {
      // Erreur non bloquante
      logWarning(`[${logId}] Historique non enregistré: ${e.message}`);
    }
    
    return result;
    
  } catch (error) {
    // ===== GESTION D'ERREUR =====
    logError(`sendEmailWithTracking [${logId}]`, error);
    
    // Tentative notification développeur
    try {
      notifyDeveloperOfFailure(params, error, logId);
    } catch (e) {
      logError(`Notification développeur impossible [${logId}]`, e);
    }
    
    // Enregistrement échec
    try {
      logEmailSent(params.to, params.subject, false, 0, error.message);
    } catch (e) {
      // Erreur non bloquante
    }
    
    return {
      success: false,
      message: error.message,
      error: error.toString(),
      timestamp: new Date().toISOString(),
      logId: logId
    };
  }
}

/**
 * ============================================
 * NOTIFICATION AUTOMATIQUE DÉVELOPPEUR
 * ============================================
 */

/**
 * 🚨 Notifie le développeur en cas d'échec d'envoi
 * 
 * @param {Object} originalParams - Paramètres email original
 * @param {Error} error - Erreur survenue
 * @param {string} logId - ID du log
 */
function notifyDeveloperOfFailure(originalParams, error, logId) {
  try {
    const devEmail = DEV_CONFIG.email || getDefaultSenderEmail();
    
    if (!devEmail || !isValidEmail(devEmail)) {
      logWarning("Email développeur invalide - notification impossible");
      return;
    }
    
    const subject = `🚨 NovaMail - Échec envoi email [${logId}]`;
    
    const htmlBody = `
      <h2>⚠️ Échec d'envoi email</h2>
      <p>Un email n'a pas pu être envoyé. Détails ci-dessous :</p>
      
      <h3>Informations de l'erreur</h3>
      <ul>
        <li><strong>Date/Heure:</strong> ${new Date().toLocaleString("fr-FR")}</li>
        <li><strong>Log ID:</strong> ${logId}</li>
        <li><strong>Message erreur:</strong> ${error.message}</li>
        <li><strong>Stack trace:</strong> <pre>${error.stack || "Non disponible"}</pre></li>
      </ul>
      
      <h3>Email concerné</h3>
      <ul>
        <li><strong>Destinataire:</strong> ${originalParams.to}</li>
        <li><strong>Sujet:</strong> ${originalParams.subject}</li>
        <li><strong>Expéditeur:</strong> ${originalParams.from || getDefaultSenderEmail()}</li>
      </ul>
      
      <h3>Actions recommandées</h3>
      <ol>
        <li>Vérifier les permissions Gmail du script</li>
        <li>Vérifier que l'email expéditeur est autorisé</li>
        <li>Consulter les logs : <code>Logger.log()</code> dans Apps Script</li>
        <li>Vérifier les quotas Gmail restants</li>
      </ol>
      
      <hr>
      <p style="font-size:12px; color:#666;">
        Cet email est généré automatiquement par NovaMail.<br>
        Projet: ${ScriptApp.getScriptId()}<br>
        Timezone: ${Session.getScriptTimeZone()}
      </p>
    `;
    
    // Envoi direct (pas de récursion - utilise GmailApp directement)
    GmailApp.sendEmail(
      devEmail,
      subject,
      stripHtml(htmlBody),
      {
        htmlBody: htmlBody,
        name: "NovaMail Error Reporter"
      }
    );
    
    logInfo(`🚨 Notification développeur envoyée à ${devEmail}`);
    
  } catch (notifError) {
    logError("notifyDeveloperOfFailure", notifError);
  }
}

/**
 * ============================================
 * LOGGING HISTORIQUE EMAILS
 * ============================================
 */

/**
 * 📊 Enregistre l'envoi d'un email dans l'historique
 * 
 * @param {string} to - Destinataire
 * @param {string} subject - Sujet
 * @param {boolean} success - Succès ou échec
 * @param {number} duration - Durée en ms
 * @param {string} errorMessage - Message d'erreur (si échec)
 */
function logEmailSent(to, subject, success, duration, errorMessage) {
  try {
    const log = loadScriptProperty("EMAIL_SEND_LOG") || [];
    
    log.push({
      to: to,
      subject: subject,
      success: success,
      duration: duration,
      error: errorMessage || null,
      timestamp: new Date().toISOString()
    });
    
    // Garder seulement les 200 derniers
    if (log.length > 200) {
      log.shift();
    }
    
    saveScriptProperty("EMAIL_SEND_LOG", log);
    
  } catch (error) {
    // Erreur silencieuse (historique optionnel)
    logWarning("logEmailSent: " + error.message);
  }
}

/**
 * Récupère l'historique des envois d'emails
 * 
 * @returns {Array<Object>} Historique
 */
function getEmailSendHistory() {
  return loadScriptProperty("EMAIL_SEND_LOG") || [];
}

/**
 * ============================================
 * WRAPPER ENVOI EMAIL BIENVENUE
 * ============================================
 */

/**
 * 🎉 Envoie l'email de bienvenue à un nouveau client
 * VERSION ROBUSTE avec tracking complet
 * 
 * @param {Client} client - Objet client
 * @returns {Object} Résultat {success, message}
 */
function sendWelcomeEmail(client) {
  try {
    if (!client || !client.loginEmail) {
      throw new Error("Client ou email manquant");
    }
    
    logInfo(`🎉 Préparation email de bienvenue pour ${client.loginEmail}...`);
    
    const config = getVersionConfig();
    const subject = "🎉 Bienvenue sur NovaMail - Votre espace est prêt !";
    const htmlBody = buildWelcomeEmailHTML(client, config);
    
    // Envoi via système robuste
    const result = sendEmailWithTracking({
      to: client.loginEmail,
      subject: subject,
      htmlBody: htmlBody,
      from: getDefaultSenderEmail(),
      fromName: "NovaMail - Équipe d'activation",
      replyTo: getDefaultSenderEmail()
    });
    
    if (result.success) {
      logInfo(`✅ Email de bienvenue envoyé à ${client.loginEmail} (${result.duration}ms)`);
    } else {
      logError("sendWelcomeEmail", new Error(result.message));
    }
    
    return result;
    
  } catch (error) {
    logError("sendWelcomeEmail", error);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * Construction du HTML de l'email de bienvenue
 * (INCHANGÉ - copié depuis UserManagement.gs)
 */
function buildWelcomeEmailHTML(client, config) {
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .header { background: linear-gradient(135deg, #1a73e8 0%, #4f46e5 100%); 
              color: white; padding: 30px; text-align: center; border-radius: 8px 8px 0 0; }
    .content { background: #f9fafb; padding: 30px; border-radius: 0 0 8px 8px; }
    .button { display: inline-block; background: #1a73e8; color: white !important; 
              padding: 14px 28px; text-decoration: none; border-radius: 6px; 
              font-weight: bold; margin: 20px 0; }
    .info-box { background: white; padding: 15px; border-left: 4px solid #1a73e8; 
                margin: 20px 0; border-radius: 4px; }
    .footer { text-align: center; margin-top: 30px; color: #6b7280; font-size: 12px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>🎉 Bienvenue sur NovaMail !</h1>
    </div>
    
    <div class="content">
      <p>Bonjour <strong>${client.fullName}</strong>,</p>
      
      <p>Merci d'avoir rejoint <strong>NovaMail</strong> ! 🚀</p>
      
      <p>Votre espace personnel est maintenant actif et prêt à l'emploi.</p>
      
      <div class="info-box">
        <h3>📊 Informations de votre compte :</h3>
        <ul>
          <li><strong>Version activée :</strong> ${config.displayName}</li>
          <li><strong>Email de connexion :</strong> ${client.loginEmail}</li>
          <li><strong>Entreprise :</strong> ${client.companyName || "Non renseignée"}</li>
          <li><strong>Quota mensuel :</strong> ${config.monthlyQuota} destinataires</li>
          <li><strong>Date d'activation :</strong> ${formatDateFR(new Date(client.activatedAt))}</li>
        </ul>
      </div>
      
      <div style="text-align: center;">
        <a href="${client.personalLink}" class="button">
          🚀 Accéder à mon espace NovaMail
        </a>
      </div>
      
      <div class="info-box">
        <h3>✨ Vos avantages ${config.name} :</h3>
        <ul>
          <li>✅ Jusqu'à <strong>${config.maxRecipients}</strong> destinataires par campagne</li>
          ${config.allowAttachments ? '<li>✅ Pièces jointes autorisées</li>' : ''}
          ${config.allowImportSheets ? '<li>✅ Import depuis Google Sheets</li>' : ''}
          ${config.allowTemplateSave ? '<li>✅ Sauvegarde de modèles d\'emails</li>' : ''}
          ${config.scheduleSend ? '<li>✅ Planification de campagnes</li>' : ''}
          ${config.customBranding ? '<li>✅ Branding personnalisé</li>' : ''}
        </ul>
      </div>
      
      <p><strong>🔗 Votre lien personnel :</strong><br>
      <a href="${client.personalLink}">${client.personalLink}</a></p>
      
      <p><em>💡 Astuce : Enregistrez ce lien dans vos favoris pour un accès rapide !</em></p>
      
      <hr style="border: none; border-top: 1px solid #e5e7eb; margin: 30px 0;">
      
      <p>Besoin d'aide ? Notre équipe est là pour vous accompagner :</p>
      <ul>
        <li>📧 Email : ${getDefaultSenderEmail()}</li>
        <li>📚 Documentation : <a href="https://docs.novamail.app">docs.novamail.app</a></li>
      </ul>
      
      <p>À très bientôt,<br>
      <strong>L'équipe NovaMail</strong> 💙</p>
    </div>
    
    <div class="footer">
      <p>© ${new Date().getFullYear()} NovaMail - Gestion de campagnes email professionnelle</p>
      <p>Vous recevez cet email suite à votre inscription sur NovaMail.</p>
    </div>
  </div>
</body>
</html>
  `.trim();
}

/**
 * ============================================
 * FONCTIONS DE DIAGNOSTIC
 * ============================================
 */

/**
 * 🧪 Teste le système d'envoi d'emails
 * 
 * @returns {Object} Résultat des tests
 */
function testEmailSystem() {
  const results = {
    success: true,
    tests: []
  };
  
  // Test 1 : DEFAULT_SENDER_EMAIL configuré
  try {
    const sender = getDefaultSenderEmail();
    results.tests.push({
      name: "Email expéditeur par défaut",
      passed: !!sender && isValidEmail(sender),
      value: sender
    });
  } catch (e) {
    results.tests.push({
      name: "Email expéditeur par défaut",
      passed: false,
      error: e.message
    });
    results.success = false;
  }
  
  // Test 2 : Permissions Gmail
  try {
    GmailApp.getAliases();
    results.tests.push({
      name: "Permissions Gmail",
      passed: true,
      value: "OK"
    });
  } catch (e) {
    results.tests.push({
      name: "Permissions Gmail",
      passed: false,
      error: e.message
    });
    results.success = false;
  }
  
  // Test 3 : Email développeur configuré
  const devEmail = DEV_CONFIG.email;
  results.tests.push({
    name: "Email développeur",
    passed: !!devEmail && isValidEmail(devEmail),
    value: devEmail || "Non configuré"
  });
  
  // Test 4 : Envoi test réel (optionnel)
  try {
    const testEmail = devEmail || getDefaultSenderEmail();
    
    const testResult = sendEmailWithTracking({
      to: testEmail,
      subject: "🧪 NovaMail - Test système email",
      htmlBody: "<p>Ce test confirme que le système d'envoi email fonctionne correctement.</p>",
      from: getDefaultSenderEmail(),
      fromName: "NovaMail Test"
    });
    
    results.tests.push({
      name: "Envoi email test",
      passed: testResult.success,
      value: testResult.message
    });
    
    if (!testResult.success) {
      results.success = false;
    }
    
  } catch (e) {
    results.tests.push({
      name: "Envoi email test",
      passed: false,
      error: e.message
    });
  }
  
  return results;
}

/**
 * 📊 Affiche les statistiques d'envoi d'emails
 * 
 * @returns {Object} Statistiques
 */
function getEmailStats() {
  const history = getEmailSendHistory();
  
  const stats = {
    total: history.length,
    successful: history.filter(e => e.success).length,
    failed: history.filter(e => !e.success).length,
    avgDuration: 0,
    recentErrors: []
  };
  
  // Calcul durée moyenne
  const successfulWithDuration = history.filter(e => e.success && e.duration);
  if (successfulWithDuration.length > 0) {
    const totalDuration = successfulWithDuration.reduce((sum, e) => sum + e.duration, 0);
    stats.avgDuration = Math.round(totalDuration / successfulWithDuration.length);
  }
  
  // Dernières erreurs (max 5)
  stats.recentErrors = history
    .filter(e => !e.success && e.error)
    .slice(-5)
    .map(e => ({
      to: e.to,
      subject: e.subject,
      error: e.error,
      timestamp: e.timestamp
    }));
  
  return stats;
}
