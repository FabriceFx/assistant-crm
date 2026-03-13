/**
 * ============================================================
 * CRM ASSISTANT GMAIL ADD-ON — Version 6.0
 * ============================================================
 */

const CONFIG = {
  GEMINI_API_KEY: 'VOTRE_CLE_API',
  GEMINI_MODEL: 'gemini-2.5-flash', 
  SPREADSHEET_ID: '', 
};

const SEARCH_QUERIES = {
  DEVIS: 'is:unread "devis" OR "cotation" OR "proforma" OR "offre commerciale"',
  COMMANDES: 'is:unread "commande" OR "validation" OR "bon de commande" OR "BC"',
  SAV: 'is:unread "réclamation" OR "litige" OR "problème" OR "SAV"'
};

const GEMINI_URL = "https://generativelanguage.googleapis.com/v1beta/models/" + CONFIG.GEMINI_MODEL + ":generateContent?key=" + CONFIG.GEMINI_API_KEY;

/**
 * Page d'accueil du module complémentaire.
 */
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
 * Affiche l'analyse IA détaillée d'un email avec une interface Material Design 3.
 */
function buildContextualCard(e) {
  try {
    const messageId = e.messageMetadata.messageId;
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

    const analysis = callGeminiAI(message.getPlainBody().substring(0, 3000), subject, pdfAttachments);

    const card = CardService.newCardBuilder();

    // En-tête de la carte
    card.setHeader(CardService.newCardHeader()
      .setTitle(String(analysis.categorie || "Dossier Client"))
      .setSubtitle("Sentiment : " + String(analysis.sentiment || "Neutre")));

    const bizSection = CardService.newCardSection().setHeader('📌 Infos Dossier');
    
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

    // Section des actions rapides avec structure Material Design 3
    const actionSection = CardService.newCardSection().setHeader('⚡ Actions rapides');
    
    // Ajout d'un séparateur visuel pour structurer l'interface
    actionSection.addWidget(CardService.newDivider());
    
    const threadUrl = "https://mail.google.com/mail/u/0/#inbox/" + thread.getId();
    
    // Premier groupe d'actions (consultation et réponse)
    const primaryButtonSet = CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText("✉️ Ouvrir l'email")
        .setOpenLink(CardService.newOpenLink().setUrl(threadUrl)));

    if (analysis.reponseSuggree) {
      primaryButtonSet.addButton(CardService.newTextButton()
        .setText('📝 Préparer le brouillon')
        // Correction effectuée ici avec setTextButtonStyle
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED) 
        .setOnClickAction(CardService.newAction()
          .setFunctionName('createReplyDraftAction')
          .setParameters({ 
            reply: String(analysis.reponseSuggree), 
            messageId: String(messageId) 
          })));
    }
    actionSection.addWidget(primaryButtonSet);

    // Deuxième groupe d'actions (organisation et CRM)
    const secondaryButtonSet = CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText('📅 Rappel (Agenda)')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('createCalendarEventAction')
          .setParameters({ 
            subject: String(subject), 
            sender: String(sender) 
          })))
      .addButton(CardService.newTextButton()
        .setText('📊 Loguer CRM')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('logToCrmAction')
          .setParameters({ 
            client: String(sender), 
            amount: String(analysis.montant || "N/A"), 
            category: String(analysis.categorie || "Autre"), 
            summary: String(analysis.resume || "") 
          })));

    actionSection.addWidget(secondaryButtonSet);
    card.addSection(actionSection);

    // Pied de page fixe avec le bouton de retour au Dashboard
    card.setFixedFooter(CardService.newFixedFooter().setPrimaryButton(
      CardService.newTextButton()
        .setText('Dashboard')
        .setOnClickAction(CardService.newAction().setFunctionName('goToHomepageAction'))
    ));

    return card.build();

  } catch (err) {
    return buildErrorCard("Erreur Analyse : " + err.message);
  }
}

/**
 * Appelle l'API Gemini pour analyser le contenu (Mail + PDF).
 */
function callGeminiAI(text, subject, pdfAttachments = []) {
  let parts = [
    { text: "Tu es un assistant Commercial CRM. Analyse l'email et les PDF joints pour extraire les montants (ex: 59,98 €), références et intentions." },
    { text: "Réponds UNIQUEMENT en JSON valide : {\"urgence\":\"...\", \"sentiment\":\"Emoji + texte\", \"categorie\":\"...\", \"montant\":\"...\", \"numeroCommande\":\"...\", \"resume\":\"...\", \"reponseSuggree\":\"...\"}" },
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
    response = UrlFetchApp.fetch(GEMINI_URL, options);
    if (response.getResponseCode() !== 429) break; 
    Utilities.sleep(Math.pow(2, i) * 1000);
  }

  try {
    const resText = response.getContentText();
    const json = JSON.parse(resText);
    if (json.candidates && json.candidates[0]) {
      let raw = json.candidates[0].content.parts[0].text;
      raw = raw.replace(/```json|```/g, "").trim();
      const parsed = JSON.parse(raw);
      // Sécurité pour s'assurer qu'aucun champ n'est null
      Object.keys(parsed).forEach(key => { if (parsed[key] === null) parsed[key] = ""; });
      return parsed;
    }
    throw new Error("IA muette");
  } catch (err) {
    return { urgence: "Moyenne", sentiment: "Neutre", categorie: "Autre", montant: "N/A", resume: "Analyse indisponible.", reponseSuggree: "" };
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
/**
 * Liste les fils de discussion pour une catégorie donnée (interface optimisée).
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
    const now = new Date();
    const startTime = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1, 9, 0, 0);
    const endTime = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1, 9, 30, 0);
    CalendarApp.getDefaultCalendar().createEvent(`[CRM] Relance : ${e.parameters.sender}`, startTime, endTime, { description: `Sujet : ${e.parameters.subject}` });
    return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("📅 Événement ajouté à demain 9h !")).build();
  } catch (err) { return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("❌ Erreur Agenda")).build(); }
}

function logToCrmAction(e) {
  try {
    const p = e.parameters;
    const ss = CONFIG.SPREADSHEET_ID ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID) : SpreadsheetApp.create("Base_CRM_Assistant");
    const sheet = ss.getSheets()[0];
    if (sheet.getLastRow() === 0) sheet.appendRow(["Date", "Client", "Catégorie", "Montant", "Résumé"]);
    sheet.appendRow([new Date(), p.client, p.category, p.amount, p.summary]);
    return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("✅ Client logué au CRM !")).build();
  } catch (err) { return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("❌ Erreur Sheets")).build(); }
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
 * Récupération des statistiques.
 */
function getDashboardStats() {
  return {
    quotes: GmailApp.search(SEARCH_QUERIES.DEVIS, 0, 50).length,
    orders: GmailApp.search(SEARCH_QUERIES.COMMANDES, 0, 50).length,
    complaints: GmailApp.search(SEARCH_QUERIES.SAV, 0, 50).length
  };
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
