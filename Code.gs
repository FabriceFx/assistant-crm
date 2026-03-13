/**
 * ============================================================
 * CRM ASSISTANT GMAIL ADD-ON
 * ============================================================
 */

const CONFIG = {
  GEMINI_MODEL: 'gemini-2.5-flash'
};

const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

const SEARCH_QUERIES = {
  DEVIS: 'is:unread "devis" OR "cotation" OR "proforma" OR "offre commerciale"',
  COMMANDES: 'is:unread "commande" OR "validation" OR "bon de commande" OR "BC"',
  SAV: 'is:unread "réclamation" OR "litige" OR "problème" OR "SAV"'
};

const GEMINI_URL = "https://generativelanguage.googleapis.com/v1beta/models/" + CONFIG.GEMINI_MODEL + ":generateContent?key=" + CONFIG.GEMINI_API_KEY;

/**
 * Page d'accueil du module complémentaire.
 */
function buildHomepage(e) {
  try {
    const card = CardService.newCardBuilder();
    
    card.setHeader(
      CardService.newCardHeader()
        .setTitle('⚡ Assistant CRM')
        // Pense à remplacer cette URL par celle de ton icône compatible mode sombre
        .setImageUrl('https://www.gstatic.com/images/icons/material/system_gm/1x/mail_black_24dp.png')
    );

    const stats = getDashboardStats();
    const statsSection = CardService.newCardSection().setHeader('📊 État de vos flux');
    
    // Remplacement des widgets avec boutons par des lignes entièrement cliquables
    statsSection.addWidget(createClickableMetric('Demandes de devis', stats.quotes, SEARCH_QUERIES.DEVIS, CardService.Icon.DESCRIPTION));
    statsSection.addWidget(createClickableMetric('Commandes / ADV', stats.orders, SEARCH_QUERIES.COMMANDES, CardService.Icon.SHOPPING_CART));
    statsSection.addWidget(createClickableMetric('SAV / Litiges', stats.complaints, SEARCH_QUERIES.SAV, CardService.Icon.TICKET));

    card.addSection(statsSection);

    const prioritySection = CardService.newCardSection().setHeader('🔴 À traiter en priorité');
    const urgentThreads = getUrgentThreads();

    if (urgentThreads.length === 0) {
      prioritySection.addWidget(CardService.newTextParagraph().setText('✅ Aucune urgence non lue.'));
    } else {
      urgentThreads.forEach((item, index) => {
        prioritySection.addWidget(
          CardService.newDecoratedText()
            .setText(String(item.subject).substring(0, 45) + "...")
            .setBottomLabel(String(item.from))
            .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.CLOCK))
            .setOnClickAction(CardService.newAction()
              .setFunctionName('viewThreadAnalysisAction')
              .setParameters({ threadId: String(item.threadId) }))
        );
        
        // Ajout d'un séparateur visuel entre les emails, sauf après le dernier
        if (index < urgentThreads.length - 1) {
          prioritySection.addWidget(CardService.newDivider());
        }
      });
    }
    card.addSection(prioritySection);

    return card.build();
  } catch (err) {
    return buildErrorCard("Erreur Dashboard : " + err.message);
  }
}

/**
 * Crée une ligne de statistique épurée et entièrement cliquable.
 * Remplace l'ancienne fonction createMetricWidget.
 */
function createClickableMetric(label, count, query, icon) {
  return CardService.newDecoratedText()
    .setTopLabel(label)
    // Mise en gras du compteur pour une meilleure hiérarchie visuelle
    .setText(`<b>${String(count)}</b> en attente`)
    .setStartIcon(CardService.newIconImage().setIcon(icon))
    // L'action est placée directement sur l'élément textuel au lieu d'un bouton externe
    .setOnClickAction(CardService.newAction()
      .setFunctionName('listCategoryThreadsAction')
      .setParameters({ category: String(label), query: String(query) }));
}

/**
 * Affiche l'analyse IA détaillée d'un email.
 */
/**
 * 1. Affiche instantanément une carte d'attente (sans bloquer l'interface).
 */
function buildContextualCard(e) {
  try {
    const messageId = e.messageMetadata.messageId;
    const message = GmailApp.getMessageById(messageId);
    const sender = message.getFrom().split('<')[0].trim();

    const card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle('Assistant IA prêt')
        .setSubtitle(sender)
      );

    const section = CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText("Cliquez ci-dessous pour extraire les données, analyser le sentiment et préparer un plan d'action avec Gemini."));

    // Action avec indicateur de chargement (SPINNER) natif
    const analyzeAction = CardService.newAction()
      .setFunctionName('performAnalysisAction')
      .setParameters({ messageId: String(messageId) })
      .setLoadIndicator(CardService.LoadIndicator.SPINNER);

    section.addWidget(CardService.newTextButton()
      .setText('✨ Lancer l\'analyse')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(analyzeAction)
    );

    card.addSection(section);
    return card.build();

  } catch (err) {
    return buildErrorCard("Erreur de préparation : " + err.message);
  }
}

/**
 * 2. Exécute l'analyse IA et met à jour l'interface avec les résultats.
 */
/**
 * 2. Exécute l'analyse IA et met à jour l'interface avec les résultats.
 */
function performAnalysisAction(e) {
  try {
    const messageId = e.parameters.messageId;
    const message = GmailApp.getMessageById(messageId);
    const thread = message.getThread();
    const subject = message.getSubject();
    const sender = message.getFrom();
    
    const attachments = message.getAttachments();
    const pdfAttachments = attachments.filter(attr => {
      const type = attr.getContentType().toLowerCase();
      const name = attr.getName().toLowerCase();
      return type.indexOf('pdf') !== -1 || name.indexOf('.pdf') !== -1;
    });

    // Appel à l'IA 
    const analysis = callGeminiAI(message.getPlainBody().substring(0, 3000), subject, pdfAttachments);

    const card = CardService.newCardBuilder();
    card.setHeader(CardService.newCardHeader()
      .setTitle(String(analysis.categorie || "Dossier Client"))
      .setSubtitle(String(sender.split('<')[0].trim())) 
    );

    // Section Indicateurs Visuels
    const indicateurSection = CardService.newCardSection();
    const urgenceRaw = String(analysis.urgence || "Normale").toLowerCase();
    let urgenceHtml = "<font color='#188038'><b>🟢 Priorité normale</b></font>";
    if (urgenceRaw.includes("haute") || urgenceRaw.includes("urgent") || urgenceRaw.includes("critique")) {
      urgenceHtml = "<font color='#d93025'><b>🔴 Priorité haute</b></font>";
    } else if (urgenceRaw.includes("moyenne")) {
      urgenceHtml = "<font color='#e37400'><b>🟠 Priorité moyenne</b></font>";
    }

    indicateurSection.addWidget(CardService.newDecoratedText().setText(urgenceHtml));
    indicateurSection.addWidget(CardService.newDecoratedText()
      .setText(`<b>${String(analysis.sentiment || "Neutre")}</b>`)
      .setBottomLabel("Sentiment détecté")
    );
    card.addSection(indicateurSection);

    // Section Infos Dossier
    const bizSection = CardService.newCardSection().setHeader('📌 Détails du dossier');
    if (pdfAttachments.length > 0) {
      bizSection.addWidget(CardService.newDecoratedText()
        .setText(`${pdfAttachments.length} PDF analysé(s)`)
        .setBottomLabel("Lecture du contenu des pièces jointes effectuée")
        .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.DESCRIPTION)));
    }

    if (analysis.montant && analysis.montant !== 'N/A') {
      bizSection.addWidget(CardService.newDecoratedText()
        .setTopLabel('Potentiel estimé')
        .setText(String(analysis.montant))
        .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.DOLLAR)));
    }
    
    bizSection.addWidget(CardService.newTextParagraph().setText("**Résumé :**\n" + String(analysis.resume || "Analyse en cours...")));
    card.addSection(bizSection);

    const threadUrl = "https://mail.google.com/mail/u/0/#inbox/" + thread.getId();

    // Section Plan d'action
    if (analysis.tacheTitre) {
      const taskSection = CardService.newCardSection().setHeader('📋 Action recommandée');
      taskSection.addWidget(CardService.newDecoratedText()
        .setText(String(analysis.tacheTitre))
        .setBottomLabel(String(analysis.tacheDescription || "Pas de détails additionnels."))
        .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.CLOCK))
      );
      
      // Tronquer la description pour éviter le dépassement de limite de paramètres Google
      const taskNotes = `${analysis.tacheDescription || ''}\n\nClient : ${sender}\nLien : ${threadUrl}`.substring(0, 1500);
      
      taskSection.addWidget(CardService.newTextButton()
        .setText('✅ Ajouter à mes tâches')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction()
          .setFunctionName('createTaskAction')
          .setParameters({ title: String(analysis.tacheTitre), notes: String(taskNotes) })));
      card.addSection(taskSection);
    }

    // Section Actions rapides
    const actionSection = CardService.newCardSection().setHeader('⚡ Actions rapides');
    actionSection.addWidget(CardService.newDivider());
    
    const primaryButtonSet = CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText("✉️ Ouvrir l'email")
        .setOpenLink(CardService.newOpenLink().setUrl(threadUrl)));

    if (analysis.reponseSuggree) {
      // SOLUTION : Stockage en cache au lieu de le passer en paramètre du bouton
      CacheService.getUserCache().put('draft_' + messageId, String(analysis.reponseSuggree), 1800); // Valable 30 minutes
      
      primaryButtonSet.addButton(CardService.newTextButton()
        .setText('📝 Préparer le brouillon')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED) 
        .setOnClickAction(CardService.newAction()
          .setFunctionName('createReplyDraftAction')
          .setParameters({ messageId: String(messageId) }))); // On ne passe QUE l'ID !
    }
    actionSection.addWidget(primaryButtonSet);

    const secondaryButtonSet = CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText('📅 Rappel (Agenda)')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('showDatePickerCardAction')
          .setParameters({ subject: String(subject), sender: String(sender) })));

    actionSection.addWidget(secondaryButtonSet);
    card.addSection(actionSection);

    card.setFixedFooter(CardService.newFixedFooter().setPrimaryButton(
      CardService.newTextButton()
        .setText('Tableau de bord')
        .setOnClickAction(CardService.newAction().setFunctionName('goToHomepageAction'))
    ));

    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(card.build()))
      .build();

  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(buildErrorCard("Erreur Analyse : " + err.message)))
      .build();
  }
}

/**
 * Appelle l'API Gemini pour analyser le contenu
 */

function callGeminiAI(text, subject, pdfAttachments = []) {
  // Récupération de la clé API au moment de l'exécution
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  
  if (!apiKey) {
    throw new Error("La clé API GEMINI_API_KEY est introuvable dans les propriétés du script.");
  }

  const modelUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey;

  let parts = [
    { text: "Tu es un assistant commercial CRM. Analyse l'email et les PDF joints pour extraire les montants, références, intentions et identifier l'action principale à réaliser." },
    { text: "Réponds UNIQUEMENT en JSON valide : {\"urgence\":\"...\", \"sentiment\":\"Emoji + texte\", \"categorie\":\"...\", \"montant\":\"...\", \"numeroCommande\":\"...\", \"resume\":\"...\", \"reponseSuggree\":\"...\", \"tacheTitre\":\"Titre de l'action à faire (court)\", \"tacheDescription\":\"Détail de l'action (court)\"}" },
    { text: `Objet : ${subject}` },
    { text: `Corps du mail : ${text}` }
  ];

  pdfAttachments.forEach(pdf => {
    try {
      parts.push({ 
        inline_data: { 
          mime_type: "application/pdf", 
          data: Utilities.base64Encode(pdf.getBytes()) 
        } 
      });
    } catch (e) { Logger.log("Erreur PDF: " + e.message); }
  });

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ 
      contents: [{ parts: parts }], 
      generationConfig: { temperature: 0.1, responseMimeType: "application/json" } 
    }),
    muteHttpExceptions: true
  };

  let response;
  for (let i = 0; i < 3; i++) {
    response = UrlFetchApp.fetch(modelUrl, options);
    if (response.getResponseCode() !== 429) break; 
    Utilities.sleep(Math.pow(2, i) * 1000);
  }

  try {
    const resText = response.getContentText();
    const json = JSON.parse(resText);

    // Interception des véritables erreurs de l'API Gemini
    if (json.error) {
      throw new Error("Refus de l'API : " + json.error.message);
    }

    if (json.candidates && json.candidates[0]) {
      let raw = json.candidates[0].content.parts[0].text;
      raw = raw.replace(/```json|```/g, "").trim();
      const parsed = JSON.parse(raw);
      Object.keys(parsed).forEach(key => { if (parsed[key] === null) parsed[key] = ""; });
      return parsed;
    }
    
    throw new Error("Format de réponse inconnu.");
  } catch (err) {
    return { 
      urgence: "Erreur", 
      sentiment: "Erreur", 
      categorie: "Erreur", 
      montant: "N/A", 
      resume: err.message, 
      reponseSuggree: "" 
    };
  }
}


/**
 * Retourne au tableau de bord.
 */
function goToHomepageAction(e) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildHomepage(e)))
    .build();
}

/**
 * Liste les fils de discussion pour une catégorie donnée.
 */

function listCategoryThreadsAction(e) {
  const categoryName = e.parameters.category || "Dossiers";
  const query = e.parameters.query || "is:unread";
  
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('📁 ' + categoryName));
    
  const section = CardService.newCardSection();
  const threads = GmailApp.search(query, 0, 20);
  
  if (!threads || threads.length === 0) {
    section.addWidget(CardService.newTextParagraph().setText("✅ Aucun message en attente dans cette catégorie."));
  } else {
    threads.forEach((t, index) => {
      // Nettoyage du nom de l'expéditeur pour un affichage plus propre
      let rawSender = t.getMessages()[0].getFrom();
      let cleanSender = rawSender.split('<')[0].replace(/"/g, '').trim();
      let subject = String(t.getFirstMessageSubject());
      
      // Troncature intelligente du sujet pour éviter qu'il ne prenne trop de lignes
      let displaySubject = subject.length > 45 ? subject.substring(0, 45) + "..." : subject;

      section.addWidget(CardService.newDecoratedText()
        .setText(displaySubject)
        .setBottomLabel(cleanSender)
        .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.EMAIL))
        // Toute la zone est cliquable pour une meilleure ergonomie
        .setOnClickAction(CardService.newAction()
          .setFunctionName('viewThreadAnalysisAction')
          .setParameters({ threadId: String(t.getId()) }))
      );
      
      // Ajout d'un séparateur discret entre les éléments, sauf pour le dernier
      if (index < threads.length - 1) {
        section.addWidget(CardService.newDivider());
      }
    });
  }
  
  card.addSection(section);
  
  // Maintien du bouton de retour fixé en bas de l'écran pour une navigation fluide
  card.setFixedFooter(CardService.newFixedFooter().setPrimaryButton(
    CardService.newTextButton()
      .setText('Retour')
      .setOnClickAction(CardService.newAction().setFunctionName('goToHomepageAction'))
  ));
  
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card.build()))
    .build();
}

/**
 * Création d'un événement d'agenda.
 */

function createCalendarEventAction(e) {
  try {
    const sender = String(e.parameters.sender || "Client");
    const subject = String(e.parameters.subject || "Sans objet");
    
    let startTimeMs = new Date().getTime() + 86400000; 
    
    if (e.formInput && e.formInput.rappelDateTime) {
      startTimeMs = Number(e.formInput.rappelDateTime.msSinceEpoch || e.formInput.rappelDateTime);
    }
    
    const startTime = new Date(startTimeMs);
    const endTime = new Date(startTime.getTime() + 30 * 60000); 

    CalendarApp.getDefaultCalendar().createEvent(
      `[CRM] Relance : ${sender}`, 
      startTime, 
      endTime, 
      { description: `Sujet : ${subject}` }
    );

    // Ajout de popCard() pour fermer la vue de sélection de date
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("📅 Événement planifié !"))
      .setNavigation(CardService.newNavigation().popCard())
      .build();
  } catch (err) { 
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("❌ Erreur Agenda : " + err.message))
      .build(); 
  }
}


function createReplyDraftAction(e) {
  try {
    const messageId = e.parameters.messageId;
    const replyText = e.parameters.reply;
    GmailApp.getMessageById(messageId).createDraftReply(replyText);
    return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("📝 Brouillon créé !")).build();
  } catch (err) { return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("❌ Erreur Brouillon")).build(); }
}

function viewThreadAnalysisAction(e) {
  const msgId = GmailApp.getThreadById(e.parameters.threadId).getMessages()[0].getId();
  return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(buildContextualCard({ messageMetadata: { messageId: String(msgId) } }))).build();
}

/**
 * Affiche une sous-carte pour sélectionner la date et l'heure du rappel.
 */
function showDatePickerCardAction(e) {
  const sender = e.parameters.sender || "Client";
  const subject = e.parameters.subject || "Sans objet";

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('📅 Planifier un rappel'));

  const section = CardService.newCardSection()
    .addWidget(CardService.newTextParagraph().setText(`Planification pour : **${sender}**`));

  const defaultDate = new Date();
  defaultDate.setDate(defaultDate.getDate() + 1);
  defaultDate.setHours(9, 0, 0, 0);

  // Le sélecteur est maintenant isolé sur cette vue
  const dateTimePicker = CardService.newDateTimePicker()
    .setTitle("Quand voulez-vous être rappelé ?")
    .setFieldName("rappelDateTime")
    .setValueInMsSinceEpoch(defaultDate.getTime());

  section.addWidget(dateTimePicker);

  // Bouton de validation qui déclenchera la création finale
  const buttonSet = CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText("Valider et planifier")
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction()
        .setFunctionName('createCalendarEventAction')
        .setParameters({ 
          subject: subject, 
          sender: sender 
        })));

  section.addWidget(buttonSet);
  card.addSection(section);

  // PushCard permet de superposer cette carte, l'utilisateur garde son contexte
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card.build()))
    .build();
}

/**
 * Récupération des statistiques.
 */
function getDashboardStats() {
  const cache = CacheService.getUserCache();
  const cached = cache.get('dashboard_stats');
  if (cached) return JSON.parse(cached);

  const stats = {
    quotes:     GmailApp.search(SEARCH_QUERIES.DEVIS, 0, 50).length,
    orders:     GmailApp.search(SEARCH_QUERIES.COMMANDES, 0, 50).length,
    complaints: GmailApp.search(SEARCH_QUERIES.SAV, 0, 50).length
  };

  cache.put('dashboard_stats', JSON.stringify(stats), 300); // 5 min
  return stats;
}

function getUrgentThreads() {
  return GmailApp.search('is:unread is:important', 0, 8).map(t => ({
    threadId: t.getId(),
    subject: t.getFirstMessageSubject(),
    from: t.getMessages()[0].getFrom().split('<')[0].replace(/"/g, '').trim()
  }));
}

function createMetricWidget(label, count, query, icon) {
  return CardService.newDecoratedText()
    .setTopLabel(label)
    .setText(String(count) + " en attente")
    .setStartIcon(CardService.newIconImage().setIcon(icon))
    .setButton(CardService.newTextButton().setText("Voir").setOnClickAction(CardService.newAction().setFunctionName('listCategoryThreadsAction').setParameters({ category: String(label), query: String(query) })));
}

/**
 * Création d'une tâche dans Google Tasks avec le lien du mail.
 */
function createTaskAction(e) {
  try {
    const title = e.parameters.title;
    const notes = e.parameters.notes;
    
    // L'API Google Tasks attend un objet spécifique
    const task = {
      title: title,
      notes: notes
    };
    
    // '@default' cible la liste de tâches principale de l'utilisateur
    Tasks.Tasks.insert(task, '@default');
    
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("✅ Tâche ajoutée à Google Tasks !"))
      .build();
  } catch (err) { 
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("❌ Erreur Tasks : " + err.message))
      .build(); 
  }
}

function buildErrorCard(message) {
  const section = CardService.newCardSection()
    .addWidget(
      CardService.newDecoratedText()
        .setTopLabel("Oups, un problème est survenu")
        .setText(String(message))
        .setWrapText(true)
    );

  const buttonSet = CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText("Retour à l'accueil")
      // Correction de la méthode ici :
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('goToHomepageAction')));
      
  section.addWidget(buttonSet);

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('❌ Interruption'))
    .addSection(section)
    .build();
}
