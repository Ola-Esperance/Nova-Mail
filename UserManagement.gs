/*****************************************************
 * NOVAMAIL SAAS - USERMANAGEMENT.GS
 * ====================================================
 * Gestion des utilisateurs et espaces clients
 * Automatisation : Tally → Google Sheets → Activation
 * 
 * @author NovaMail Team
 * @version 2.0.0
 * @lastModified 2025-11-03
 *****************************************************/

/**
 * ============================================
 * CONSTANTES MODULE
 * ============================================
 */

// Clés de stockage pour les clients
const CLIENT_STORAGE = {
  PREFIX: "CLIENT_",                    // Préfixe pour tous les clients
  INDEX_KEY: "CLIENT_INDEX",            // Index de tous les userId
  ACTIVATION_LOG: "ACTIVATION_LOG"      // Log des activations
};

// Configuration email de bienvenue
const WELCOME_EMAIL_CONFIG = {
  senderName: "NovaMail - Équipe d'activation",
  subject: "🎉 Bienvenue sur NovaMail - Votre espace est prêt !",
  replyTo: DEFAULT_SENDER_EMAIL
};

// Colonnes attendues dans Google Sheets (Tally)
const EXPECTED_COLUMNS = {
  submissionId: "Submission ID",
  respondentId: "Respondent ID",
  submittedAt: "Submitted at",
  fullName: "Nom complet",
  loginEmail: "Email de connexion",
  senderEmail: "Email d'envoi",
  replyEmail: "Email de réponse (optionnel)",
  companyName: "Nom de l'entreprise / organisation",
  activatedVersion: "Version activée",
  consent: "Consentement",
  consentValue: "Consentement (Accepter)"
};

/**
 * ============================================
 * STRUCTURE DE DONNÉES CLIENT
 * ============================================
 */

/**
 * Définition de l'objet Client
 * @typedef {Object} Client
 * @property {string} userId - Identifiant unique client (généré)
 * @property {string} submissionId - ID soumission Tally
 * @property {string} respondentId - ID répondant Tally
 * @property {string} fullName - Nom complet
 * @property {string} loginEmail - Email de connexion (unique)
 * @property {string} senderEmail - Email expéditeur campagnes
 * @property {string} replyEmail - Email de réponse (optionnel)
 * @property {string} companyName - Nom entreprise/organisation
 * @property {string} version - Version activée (FREE|STARTER|PRO|BUSINESS)
 * @property {string} activatedAt - Date activation (ISO)
 * @property {string} personalLink - Lien personnel vers espace
 * @property {boolean} emailSent - Email de bienvenue envoyé
 * @property {string} status - Statut (active|suspended|cancelled)
 * @property {Object} metadata - Métadonnées supplémentaires
 */

/**
 * ============================================
 * ACTIVATION AUTOMATIQUE DEPUIS GOOGLE SHEETS
 * ============================================
 */

/**
 * 🚀 Point d'entrée principal : traite une nouvelle ligne Google Sheets
 * 
 * Cette fonction est appelée automatiquement quand une nouvelle soumission
 * Tally arrive dans le Google Sheet.
 * 
 * @param {Object} rowData - Données de la ligne (objet clé-valeur)
 * @returns {Object} Résultat de l'activation {success, userId, message}
 * 
 * @example
 * // Appelé automatiquement par trigger ou manuellement :
 * processNewClientSubmission({
 *   "Email de connexion": "john@example.com",
 *   "Nom complet": "John Doe",
 *   "Version activée": "PRO",
 *   ...
 * });
 */
function processNewClientSubmission(rowData) {
  try {
    logInfo("🔄 Traitement nouvelle soumission client...");
    
    // 1️⃣ Validation des données reçues
    validateClientData(rowData);
    
    // 2️⃣ Vérification si client déjà existant (email unique)
    const loginEmail = normalizeEmail(rowData[EXPECTED_COLUMNS.loginEmail]);
    const existingClient = findClientByEmail(loginEmail);
    
    if (existingClient) {
      logWarning(`Client déjà activé : ${loginEmail}`);
      return {
        success: false,
        userId: existingClient.userId,
        message: "Ce client est déjà enregistré.",
        existingClient: true
      };
    }
    
    // 3️⃣ Génération userId unique
    const userId = generateClientUserId();
    
    // 4️⃣ Construction objet client
    const client = buildClientObject(userId, rowData);
    
    // 5️⃣ Génération du lien personnel sécurisé
    client.personalLink = generatePersonalLink(userId);
    
    // 6️⃣ Sauvegarde dans PropertiesService
    saveClient(client);
    
    // 7️⃣ Envoi email de bienvenue
    const emailResult = sendWelcomeEmail(client);
    client.emailSent = emailResult.success;
    
    // 8️⃣ Mise à jour après envoi email
    saveClient(client);
    
    // 9️⃣ Log dans l'historique des activations
    logActivation(client, emailResult);
    
    logInfo(`✅ Client activé avec succès : ${client.fullName} (${userId})`);
    
    return {
      success: true,
      userId: userId,
      message: "Client activé et email envoyé avec succès.",
      personalLink: client.personalLink
    };
    
  } catch (error) {
    logError("processNewClientSubmission", error);
    return {
      success: false,
      message: formatErrorForUser(error, "Activation client")
    };
  }
}

/**
 * ============================================
 * VALIDATION DES DONNÉES
 * ============================================
 */

/**
 * Valide les données reçues du formulaire Tally
 * 
 * @param {Object} rowData - Données de la ligne
 * @throws {Error} Si données invalides ou incomplètes
 * @returns {boolean} true si valide
 */
function validateClientData(rowData) {
  if (!rowData || typeof rowData !== "object") {
    throw new Error("Données de soumission invalides");
  }
  
  // Champs obligatoires
  const requiredFields = [
    EXPECTED_COLUMNS.loginEmail,
    EXPECTED_COLUMNS.fullName,
    EXPECTED_COLUMNS.activatedVersion,
    EXPECTED_COLUMNS.consent
  ];
  
  const missing = [];
  requiredFields.forEach(field => {
    if (!rowData[field] || String(rowData[field]).trim() === "") {
      missing.push(field);
    }
  });
  
  if (missing.length > 0) {
    throw new Error(`Champs obligatoires manquants : ${missing.join(", ")}`);
  }
  
  // Validation email
  const loginEmail = rowData[EXPECTED_COLUMNS.loginEmail];
  if (!isValidEmail(loginEmail)) {
    throw new Error(`Email de connexion invalide : ${loginEmail}`);
  }
  
  // Validation version
  const version = String(rowData[EXPECTED_COLUMNS.activatedVersion]).toUpperCase().trim();
  const validVersions = ["FREE", "STARTER", "PRO", "BUSINESS"];
  if (!validVersions.includes(version)) {
    throw new Error(`Version invalide : ${version}. Valeurs acceptées : ${validVersions.join(", ")}`);
  }
  
  // Validation consentement
  const consent = rowData[EXPECTED_COLUMNS.consentValue];
  if (consent !== true && consent !== "TRUE" && consent !== "Accepter") {
    throw new Error("Consentement non validé");
  }
  
  return true;
}

/**
 * ============================================
 * CONSTRUCTION OBJET CLIENT
 * ============================================
 */

/**
 * Construit un objet Client complet depuis les données brutes
 * 
 * @param {string} userId - ID unique généré
 * @param {Object} rowData - Données du formulaire
 * @returns {Client} Objet client structuré
 */
function buildClientObject(userId, rowData) {
  const now = new Date().toISOString();
  
  // Normalisation de la version
  const version = String(rowData[EXPECTED_COLUMNS.activatedVersion] || "FREE")
    .toUpperCase()
    .trim();
  
  // Construction objet
  const client = {
    // Identifiants
    userId: userId,
    submissionId: rowData[EXPECTED_COLUMNS.submissionId] || "",
    respondentId: rowData[EXPECTED_COLUMNS.respondentId] || "",
    
    // Informations personnelles
    fullName: String(rowData[EXPECTED_COLUMNS.fullName] || "").trim(),
    loginEmail: normalizeEmail(rowData[EXPECTED_COLUMNS.loginEmail]),
    senderEmail: normalizeEmail(rowData[EXPECTED_COLUMNS.senderEmail] || rowData[EXPECTED_COLUMNS.loginEmail]),
    replyEmail: normalizeEmail(rowData[EXPECTED_COLUMNS.replyEmail] || rowData[EXPECTED_COLUMNS.loginEmail]),
    companyName: String(rowData[EXPECTED_COLUMNS.companyName] || "").trim(),
    
    // Produit
    version: version,
    
    // Dates et statut
    submittedAt: rowData[EXPECTED_COLUMNS.submittedAt] || now,
    activatedAt: now,
    status: "active",
    
    // Communication
    emailSent: false,
    personalLink: "",
    
    // Métadonnées
    metadata: {
      source: "tally",
      userAgent: "",
      consentGiven: true,
      activationMethod: "automatic"
    }
  };
  
  return client;
}

/**
 * ============================================
 * GÉNÉRATION IDENTIFIANTS & LIENS
 * ============================================
 */

/**
 * Génère un userId unique et non prédictible
 * Format : CLIENT_[timestamp]_[random]
 * 
 * @returns {string} userId unique
 * 
 * @example
 * generateClientUserId() // "CLIENT_1699012345_a7k3p9"
 */
function generateClientUserId() {
  const timestamp = new Date().getTime();
  const random = generateShortId();
  return `${CLIENT_STORAGE.PREFIX}${timestamp}_${random}`;
}

/**
 * Génère le lien personnel sécurisé vers l'espace client
 * 
 * @param {string} userId - Identifiant client
 * @returns {string} URL complète avec paramètre userId
 * 
 * @example
 * generatePersonalLink("CLIENT_123_abc") 
 * // "https://script.google.com/macros/s/DEPLOYMENT_ID/exec?userId=CLIENT_123_abc"
 */
function generatePersonalLink(userId) {
  const deploymentId = getDeploymentId();
  
  if (!deploymentId) {
    logWarning("⚠️ Deployment ID non configuré - lien générique créé");
    return `https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec?userId=${encodeURIComponent(userId)}`;
  }
  
  return `https://script.google.com/macros/s/${deploymentId}/exec?userId=${encodeURIComponent(userId)}`;
}

/**
 * Récupère le Deployment ID du script (à configurer)
 * 
 * @returns {string|null} Deployment ID ou null
 */
function getDeploymentId() {
  // Méthode 1 : Stocké dans Script Properties
  const stored = PropertiesService.getScriptProperties()
    .getProperty("DEPLOYMENT_ID");
  
  if (stored) return stored;
  
  // Méthode 2 : Retour null si non configuré
  // L'admin doit définir via : setDeploymentId("AKfycbz...")
  return null;
}

/**
 * Configure le Deployment ID (à appeler une fois après déploiement)
 * 
 * @param {string} deploymentId - ID du déploiement web
 * @returns {boolean} Succès
 * 
 * @example
 * setDeploymentId("AKfycbzXXXXXXXXXXXXXXXXXXXX");
 */
function setDeploymentId(deploymentId) {
  try {
    PropertiesService.getScriptProperties()
      .setProperty("DEPLOYMENT_ID", deploymentId);
    logInfo("✅ Deployment ID configuré : " + deploymentId);
    return true;
  } catch (error) {
    logError("setDeploymentId", error);
    return false;
  }
}

/**
 * ============================================
 * STOCKAGE & RÉCUPÉRATION CLIENTS
 * ============================================
 */

/**
 * Sauvegarde un client dans PropertiesService
 * 
 * @param {Client} client - Objet client à sauvegarder
 * @returns {boolean} Succès de l'opération
 */
function saveClient(client) {
  try {
    // Sauvegarde client individuel
    const success = saveScriptProperty(client.userId, client);
    
    if (success) {
      // Mise à jour de l'index global
      updateClientIndex(client.userId, client.loginEmail);
    }
    
    return success;
  } catch (error) {
    logError("saveClient", error);
    return false;
  }
}

/**
 * Charge un client par son userId
 * 
 * @param {string} userId - Identifiant client
 * @returns {Client|null} Objet client ou null si inexistant
 */
function loadClient(userId) {
  if (!userId) return null;
  
  try {
    return loadScriptProperty(userId);
  } catch (error) {
    logError("loadClient", error);
    return null;
  }
}

/**
 * Recherche un client par email de connexion
 * 
 * @param {string} email - Email à rechercher
 * @returns {Client|null} Client trouvé ou null
 */
function findClientByEmail(email) {
  const normalizedEmail = normalizeEmail(email);
  const index = loadClientIndex();
  
  // Recherche dans l'index
  const userId = index[normalizedEmail];
  
  if (userId) {
    return loadClient(userId);
  }
  
  return null;
}

/**
 * Met à jour l'index global des clients (email → userId)
 * 
 * @param {string} userId - ID client
 * @param {string} email - Email du client
 * @returns {boolean} Succès
 */
function updateClientIndex(userId, email) {
  try {
    const index = loadClientIndex();
    index[normalizeEmail(email)] = userId;
    return saveScriptProperty(CLIENT_STORAGE.INDEX_KEY, index);
  } catch (error) {
    logError("updateClientIndex", error);
    return false;
  }
}

/**
 * Charge l'index global des clients
 * 
 * @returns {Object} Index email → userId
 */
function loadClientIndex() {
  const index = loadScriptProperty(CLIENT_STORAGE.INDEX_KEY);
  return index || {};
}

/**
 * Liste tous les clients actifs
 * 
 * @returns {Array<Client>} Tableau de clients
 */
function listAllClients() {
  try {
    const index = loadClientIndex();
    const clients = [];
    
    for (const email in index) {
      const userId = index[email];
      const client = loadClient(userId);
      if (client && client.status === "active") {
        clients.push(client);
      }
    }
    
    return clients;
  } catch (error) {
    logError("listAllClients", error);
    return [];
  }
}

/**
 * ============================================
 * EMAIL DE BIENVENUE
 * ============================================
 */

/**
 * Envoie l'email de bienvenue personnalisé au nouveau client
 * 
 * @param {Client} client - Objet client
 * @returns {Object} Résultat {success, message}
 */
function sendWelcomeEmail(client) {
  try {
    const config = getVersionConfig(); // Config de sa version
    
    // Construction du sujet personnalisé
    const subject = WELCOME_EMAIL_CONFIG.subject;
    
    // Construction du corps HTML
    const htmlBody = buildWelcomeEmailHTML(client, config);
    
    // Envoi via Gmail
    GmailApp.sendEmail(
      client.loginEmail,
      subject,
      stripHtml(htmlBody), // Version texte brut
      {
        htmlBody: htmlBody,
        name: WELCOME_EMAIL_CONFIG.senderName,
        replyTo: WELCOME_EMAIL_CONFIG.replyTo
      }
    );
    
    logInfo(`📧 Email de bienvenue envoyé à ${client.loginEmail}`);
    
    return {
      success: true,
      message: "Email envoyé avec succès"
    };
    
  } catch (error) {
    logError("sendWelcomeEmail", error);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * Construit le HTML de l'email de bienvenue
 * 
 * @param {Client} client - Client
 * @param {Object} config - Config version
 * @returns {string} HTML de l'email
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
        <li>📧 Email : ${WELCOME_EMAIL_CONFIG.replyTo}</li>
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
 * HISTORIQUE DES ACTIVATIONS
 * ============================================
 */

/**
 * Enregistre une activation dans l'historique
 * 
 * @param {Client} client - Client activé
 * @param {Object} emailResult - Résultat envoi email
 */
function logActivation(client, emailResult) {
  try {
    const log = loadScriptProperty(CLIENT_STORAGE.ACTIVATION_LOG) || [];
    
    log.push({
      userId: client.userId,
      email: client.loginEmail,
      fullName: client.fullName,
      version: client.version,
      activatedAt: client.activatedAt,
      emailSent: emailResult.success,
      status: "completed"
    });
    
    // Garder seulement les 500 dernières activations
    if (log.length > 500) {
      log.shift();
    }
    
    saveScriptProperty(CLIENT_STORAGE.ACTIVATION_LOG, log);
    
  } catch (error) {
    logError("logActivation", error);
  }
}

/**
 * Récupère l'historique des activations
 * 
 * @returns {Array} Historique des activations
 */
function getActivationHistory() {
  return loadScriptProperty(CLIENT_STORAGE.ACTIVATION_LOG) || [];
}

/**
 * ============================================
 * GESTION DU DOGET (POINT D'ENTRÉE WEB)
 * ============================================
 */

/**
 * 🌐 Point d'entrée HTTP pour l'accès aux espaces personnels
 * Gère les requêtes GET avec paramètre userId
 * 
 * @param {Object} e - Event object (contient e.parameter.userId)
 * @returns {HtmlOutput} Page HTML personnalisée
 * 
 * @example
 * URL : https://script.google.com/macros/s/DEPLOYMENT_ID/exec?userId=CLIENT_123_abc
 */
function doGet(e) {
  try {
    // Récupération du userId depuis l'URL
    const userId = e.parameter.userId;
    
    // Si pas de userId : afficher page d'accueil générique
    if (!userId) {
      return HtmlService.createHtmlOutputFromFile("index")
        .setTitle("NovaMail - Gestion de campagnes email");
    }
    
    // Chargement du client
    const client = loadClient(userId);
    
    // Si client inexistant : erreur 404
    if (!client) {
      return HtmlService.createHtmlOutput(`
        <h1>❌ Espace introuvable</h1>
        <p>Le lien que vous avez utilisé est invalide ou expiré.</p>
        <p>Contactez-nous si le problème persiste : ${DEFAULT_SENDER_EMAIL}</p>
      `).setTitle("Erreur - NovaMail");
    }
    
    // Si client suspendu/annulé
    if (client.status !== "active") {
      return HtmlService.createHtmlOutput(`
        <h1>⚠️ Compte suspendu</h1>
        <p>Votre compte NovaMail est actuellement ${client.status}.</p>
        <p>Pour plus d'informations : ${DEFAULT_SENDER_EMAIL}</p>
      `).setTitle("Compte suspendu - NovaMail");
    }
    
    // ✅ Client valide : injecter sa version et charger l'interface
    return loadPersonalizedInterface(client);
    
  } catch (error) {
    logError("doGet", error);
    return HtmlService.createHtmlOutput(`
      <h1>❌ Erreur serveur</h1>
      <p>Une erreur est survenue. Veuillez réessayer.</p>
      <p>Code erreur : ${error.message}</p>
    `).setTitle("Erreur - NovaMail");
  }
}

/**
 * Charge l'interface personnalisée pour un client
 * 
 * @param {Client} client - Client authentifié
 * @returns {HtmlOutput} Interface HTML personnalisée
 */
function loadPersonalizedInterface(client) {
  // Configuration de la version utilisateur dans la session
  setUserVersion(client.version);
  
  // Chargement du template HTML standard
  const template = HtmlService.createTemplateFromFile("index");
  
  // Injection des données client dans le template
  template.clientData = {
    userId: client.userId,
    fullName: client.fullName,
    email: client.loginEmail,
    company: client.companyName,
    version: client.version
  };
  
  // Évaluation et retour
  return template.evaluate()
    .setTitle(`NovaMail - ${client.fullName}`)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}
