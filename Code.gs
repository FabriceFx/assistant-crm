/**
 * ============================================================================
 * NOM DU PROJET   : Mon assistant Gmail Add-on avec IA (Gemini)
 * DESCRIPTION     : Module complémentaire Google Workspace pour Gmail. 
 * Ce script permet d'analyser, catégoriser et traiter les 
 * emails depuis la boîte de réception. Il intègre 
 * un tableau de bord de suivi, des actions rapides (tâches, 
 * agenda, brouillons) et une analyse approfondie des messages 
 * et de leurs pièces jointes (PDF) propulsée par l'IA Gemini.
 * * AUTEUR          : Fabrice FAUCHEUX
 * VERSION         : 1.1.0 
 * DATE            : Mars 2026
 * SERVICES GOOGLE : GmailApp, CardService, PropertiesService, CacheService, 
 * CalendarApp, Tasks, UrlFetchApp.
 * PRÉREQUIS D'INSTALLATION :
 * 1. Ajouter une clé API valide dans les Propriétés du script :
 * -> Nom de la propriété : GEMINI_API_KEY
 * -> Valeur : [Ta clé API Gemini]
 * 2. S'assurer que les services avancés "Tasks API" et "Calendar API" 
 * sont bien activés dans l'éditeur Apps Script si nécessaire.
 * ============================================================================
 */

/**
 * Paramètres par défaut (Noms, Requêtes, Icônes)
 */
const DEFAULT_SETTINGS = {
  LABEL_CAT1: 'Emails non lus',
  QUERY_CAT1: 'is:unread',
  ICON_CAT1: 'INVITE',

  LABEL_CAT2: 'Emails de la semaine',
  QUERY_CAT2: 'newer_than:7d',
  ICON_CAT2: 'CLOCK',

  LABEL_CAT3: 'À rappeler',
  QUERY_CAT3: 'is:unread ("appel" OR "téléphone" OR "téléphonique" OR "rappeler" OR "call" OR "vive voix")',
  ICON_CAT3: 'PHONE'
};

const AVAILABLE_ICONS = [
  // Existantes
  { label: "📄 Document / Devis",   value: "DESCRIPTION"   },
  { label: "🛒 Panier / Commande",  value: "SHOPPING_CART"  },
  { label: "🎫 Ticket / SAV",       value: "TICKET"         },
  { label: "👤 Personne / Client",  value: "PERSON"         },
  { label: "👥 Groupe / Équipe",    value: "MULTIPLE_PEOPLE" },
  { label: "💲 Finance / Argent",   value: "DOLLAR"         },
  { label: "⭐ Étoile / Important", value: "STAR"           },
  { label: "✉️ Email / Enveloppe",  value: "EMAIL"          },
  { label: "🕒 Horloge / Attente",  value: "CLOCK"          },
  { label: "📞 Téléphone / Appel",  value: "PHONE"          },
  { label: "🏬 Boutique / Magasin", value: "STORE"          },
  { label: "🏷️ Offre / Promotion",  value: "OFFER"          },
  { label: "📍 Localisation",       value: "MAP_PIN"        },
  { label: "🎟️ Réservation / N°",   value: "CONFIRMATION_NUMBER_ICON" },
  { label: "🎭 Intervenant",        value: "EVENT_PERFORMER" },
  { label: "🪑 Place / Siège",      value: "EVENT_SEAT"     },
  { label: "🎬 Vidéo / Démo",       value: "VIDEO_CAMERA"   },
  { label: "▶️ Lecture / Tuto",     value: "VIDEO_PLAY"     },
  { label: "🔖 Signet / Favori",    value: "BOOKMARK"       },
  { label: "🚂 Livraison / Train",  value: "TRAIN"          },
  { label: "✈️ Déplacement / Avion",value: "AIRPLANE"       },
  { label: "🏨 Hôtel / Hébergement",value: "HOTEL"          },
  { label: "🍽️ Restaurant / RDV",   value: "RESTAURANT_ICON" },
  { label: "🎫 Adhésion / Abonnt.", value: "MEMBERSHIP"     },
  { label: "📩 Invitation",         value: "INVITE"         }
];

// ============================================================
// PARAMÈTRES UTILISATEUR
// ============================================================

/**
 * Récupère les paramètres personnalisés ou les valeurs par défaut.
 */
function getUserSettings() {
  const userProps = PropertiesService.getUserProperties();
  return {
    LABEL_CAT1: userProps.getProperty('LABEL_CAT1') || DEFAULT_SETTINGS.LABEL_CAT1,
    QUERY_CAT1: userProps.getProperty('QUERY_CAT1') || DEFAULT_SETTINGS.QUERY_CAT1,
    ICON_CAT1:  userProps.getProperty('ICON_CAT1')  || DEFAULT_SETTINGS.ICON_CAT1,

    LABEL_CAT2: userProps.getProperty('LABEL_CAT2') || DEFAULT_SETTINGS.LABEL_CAT2,
    QUERY_CAT2: userProps.getProperty('QUERY_CAT2') || DEFAULT_SETTINGS.QUERY_CAT2,
    ICON_CAT2:  userProps.getProperty('ICON_CAT2')  || DEFAULT_SETTINGS.ICON_CAT2,

    LABEL_CAT3: userProps.getProperty('LABEL_CAT3') || DEFAULT_SETTINGS.LABEL_CAT3,
    QUERY_CAT3: userProps.getProperty('QUERY_CAT3') || DEFAULT_SETTINGS.QUERY_CAT3,
    ICON_CAT3:  userProps.getProperty('ICON_CAT3')  || DEFAULT_SETTINGS.ICON_CAT3
  };
}

/**
 * Récupère les requêtes personnalisées de l'utilisateur ou les valeurs par défaut.
 */
function getUserSearchQueries() {
  const userProps = PropertiesService.getUserProperties();
  return {
    DEVIS:     userProps.getProperty('QUERY_DEVIS')     || DEFAULT_SEARCH_QUERIES.DEVIS,
    COMMANDES: userProps.getProperty('QUERY_COMMANDES') || DEFAULT_SEARCH_QUERIES.COMMANDES,
    SAV:       userProps.getProperty('QUERY_SAV')       || DEFAULT_SEARCH_QUERIES.SAV
  };
}

// ============================================================
// TABLEAU DE BORD (HOMEPAGE)
// ============================================================

/**
 * Page d'accueil du module complémentaire.
 */
function buildHomepage(e) {
  try {
    const card = CardService.newCardBuilder();

    // Ajout d'une action dans le menu (les 3 petits points en haut à droite)
    card.addCardAction(
      CardService.newCardAction()
        .setText('⚙️ Modifier les filtres')
        .setOnClickAction(CardService.newAction().setFunctionName('buildSettingsCard'))
    );

    card.setHeader(CardService.newCardHeader());

    // Récupération dynamique des paramètres
    const settings = getUserSettings();
    const stats    = getDashboardStats();

    const statsSection = CardService.newCardSection().setHeader('📊 <b>État de votre messagerie</b>');


    const safeIcon1 = CardService.Icon[settings.ICON_CAT1] || CardService.Icon.DESCRIPTION;
    const safeIcon2 = CardService.Icon[settings.ICON_CAT2] || CardService.Icon.SHOPPING_CART;
    const safeIcon3 = CardService.Icon[settings.ICON_CAT3] || CardService.Icon.TICKET;

    statsSection.addWidget(createClickableMetric(settings.LABEL_CAT1, stats.cat1, settings.QUERY_CAT1, safeIcon1));
    statsSection.addWidget(createClickableMetric(settings.LABEL_CAT2, stats.cat2, settings.QUERY_CAT2, safeIcon2));
    statsSection.addWidget(createClickableMetric(settings.LABEL_CAT3, stats.cat3, settings.QUERY_CAT3, safeIcon3));

    card.addSection(statsSection);

    const prioritySection = CardService.newCardSection().setHeader('🔴 <b>À traiter en priorité</b>');
    const urgentThreads   = getUrgentThreads();

    if (urgentThreads.length === 0) {
      prioritySection.addWidget(
        CardService.newTextParagraph().setText('✅ Aucune urgence non lue.')
      );
    } else {
      urgentThreads.forEach((item, index) => {
        prioritySection.addWidget(
          CardService.newDecoratedText()
            .setText(String(item.subject).substring(0, 45) + "...")
            .setBottomLabel(String(item.from))
            .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.CLOCK))
            .setOnClickAction(
              CardService.newAction()
                .setFunctionName('viewThreadAnalysisAction')
                .setParameters({ threadId: String(item.threadId) })
            )
        );

        // Séparateur visuel entre les emails, sauf après le dernier
        if (index < urgentThreads.length - 1) {
          prioritySection.addWidget(CardService.newDivider());
        }
      });
    }

    card.addSection(prioritySection);

    card.setFixedFooter(
      CardService.newFixedFooter().setPrimaryButton(
        CardService.newTextButton()
          .setText('👨‍💻 Développé par Fabrice')
          .setOpenLink(CardService.newOpenLink().setUrl('https://faucheux.bzh')) 
      )
    );



    return card.build();

  } catch (err) {
    return buildErrorCard("Erreur Dashboard : " + err.message);
  }
}

/**
 * Crée une ligne de statistique épurée et entièrement cliquable.
 */
function createClickableMetric(label, count, query, icon) {
  return CardService.newDecoratedText()
    .setTopLabel(label)
    .setText(`<b>${String(count)}</b> en attente`)
    .setStartIcon(CardService.newIconImage().setIcon(icon))
    .setOnClickAction(
      CardService.newAction()
        .setFunctionName('listCategoryThreadsAction')
        .setParameters({ category: String(label), query: String(query) })
    );
}

/**
 * Widget de métrique cliquable (version legacy avec bouton externe).
 */
function createMetricWidget(label, count, query, icon) {
  return CardService.newDecoratedText()
    .setTopLabel(label)
    .setText(String(count) + " en attente")
    .setStartIcon(CardService.newIconImage().setIcon(icon))
    .setButton(
      CardService.newTextButton()
        .setText("Voir")
        .setOnClickAction(
          CardService.newAction()
            .setFunctionName('listCategoryThreadsAction')
            .setParameters({ category: String(label), query: String(query) })
        )
    );
}

// ============================================================
// ANALYSE IA
// ============================================================

/**
 * Affiche instantanément une carte d'attente avant l'analyse IA.
 */

function buildContextualCard(e) {
  try {
    const messageId = e.messageMetadata.messageId;
    const message   = GmailApp.getMessageById(messageId);
    
    // 1. Extraction des caractéristiques principales
    const rawSender   = message.getFrom();
    const senderName  = rawSender.split('<')[0].trim();
    let senderEmail   = rawSender; // Valeur par défaut si aucun chevron n'est présent
    if (rawSender.includes('<') && rawSender.includes('>')) {
      senderEmail = rawSender.split('<')[1].split('>')[0].trim();
    }
    const subject     = message.getSubject() || "Sans objet";
    const dateStr     = message.getDate().toLocaleString('fr-FR', { dateStyle: 'medium', timeStyle: 'short' });
    const attachments = message.getAttachments();
    
    // --- NOUVEAUTÉ : Génération du lien de l'email ---
    const threadId  = message.getThread().getId();
    const threadUrl = "https://mail.google.com/mail/u/0/#inbox/" + threadId;

    // 2. Création de l'en-tête de la carte
    const card = CardService.newCardBuilder()
      .setHeader(
        CardService.newCardHeader()
          .setTitle('Aperçu du message')
          .setSubtitle(senderName + " " + senderEmail)
      );

    // 3. Section : Caractéristiques de l'email
    const infoSection = CardService.newCardSection().setHeader('Caractéristiques principales');

    infoSection.addWidget(
      CardService.newDecoratedText()
        .setTopLabel("Objet")
        .setText(subject.length > 60 ? subject.substring(0, 60) + "..." : subject)
        .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.EMAIL))
    );

    infoSection.addWidget(
      CardService.newDecoratedText()
        .setTopLabel("Reçu le")
        .setText(dateStr)
        .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.CLOCK))
    );

    if (attachments.length > 0) {
      
      // On compte les PDF en vérifiant le type de contenu OU l'extension du fichier
      const pdfCount = attachments.filter(a => {
      const mimeType = a.getContentType().toLowerCase();
      const fileName = a.getName().toLowerCase();
      
      // Retourne vrai si 'pdf' est dans le type OU '.pdf' est dans le nom
      return mimeType.includes('pdf') || fileName.includes('.pdf');
      }).length;
      
      infoSection.addWidget(
        CardService.newDecoratedText()
          .setTopLabel("Pièces jointes")
          .setText(`${attachments.length} fichier(s) dont ${pdfCount} PDF`)
          .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.DESCRIPTION))
      );
    }

    infoSection.addWidget(
      CardService.newButtonSet()
        .addButton(
          CardService.newTextButton()
            .setText("↗️ Ouvrir l'email")
            .setOpenLink(CardService.newOpenLink().setUrl(threadUrl))
        )
    );

    card.addSection(infoSection);

    // 4. Section : Action IA
    const analyzeAction = CardService.newAction()
      .setFunctionName('performAnalysisAction')
      .setParameters({ messageId: String(messageId) })
      .setLoadIndicator(CardService.LoadIndicator.SPINNER);

    const aiSection = CardService.newCardSection()
      .setHeader("Assistant d'analyse")
      .addWidget(
        CardService.newTextParagraph().setText(
          "Cliquez ci-dessous pour extraire les données, analyser le sentiment " +
          "et préparer un plan d'action."
        )
      )
      .addWidget(
        CardService.newTextButton()
          .setText("Lancer l'analyse")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(analyzeAction)
      );

    card.addSection(aiSection);

    return card.build();

  } catch (err) {
    return buildErrorCard("Erreur de préparation : " + err.message);
  }
}

/**
 * Exécute l'analyse IA et met à jour l'interface avec les résultats.
 */
function performAnalysisAction(e) {
  try {
    const messageId   = e.parameters.messageId;
    const message     = GmailApp.getMessageById(messageId);
    const thread      = message.getThread();
    const subject     = message.getSubject();
    const sender      = message.getFrom();
    const attachments = message.getAttachments();

    const pdfAttachments = attachments.filter(attr => {
      const type = attr.getContentType().toLowerCase();
      const name = attr.getName().toLowerCase();
      return type.indexOf('pdf') !== -1 || name.indexOf('.pdf') !== -1;
    });

    // Appel à l'IA
    const analysis = callGeminiAI(message.getPlainBody().substring(0, 3000), subject, pdfAttachments);

    const card = CardService.newCardBuilder();
    card.setHeader(
      CardService.newCardHeader()
        .setTitle(String(analysis.categorie || "Dossier"))
        .setSubtitle(String(sender.split('<')[0].trim()))
    );

    // --- Section Indicateurs Visuels ---
    const indicateurSection = CardService.newCardSection();
    const urgenceRaw = String(analysis.urgence || "Normale").toLowerCase();
    let urgenceHtml  = "<font color='#188038'><b>🟢 Priorité normale</b></font>";

    if (urgenceRaw.includes("haute") || urgenceRaw.includes("urgent") || urgenceRaw.includes("critique")) {
      urgenceHtml = "<font color='#d93025'><b>🔴 Priorité haute</b></font>";
    } else if (urgenceRaw.includes("moyenne")) {
      urgenceHtml = "<font color='#e37400'><b>🟠 Priorité moyenne</b></font>";
    }

    indicateurSection.addWidget(CardService.newDecoratedText().setText(urgenceHtml));
    indicateurSection.addWidget(
      CardService.newDecoratedText()
        .setText(`<b>${String(analysis.sentiment || "Neutre")}</b>`)
        .setBottomLabel("Sentiment détecté")
    );
    card.addSection(indicateurSection);

    // --- Section Infos Dossier ---
    const bizSection = CardService.newCardSection().setHeader('📌 Détails du mail');

    if (pdfAttachments.length > 0) {
      bizSection.addWidget(
        CardService.newDecoratedText()
          .setText(`${pdfAttachments.length} PDF analysé(s)`)
          .setBottomLabel("Lecture du contenu des pièces jointes effectuée")
          .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.DESCRIPTION))
      );
    }

    if (analysis.montant && analysis.montant !== 'N/A') {
      bizSection.addWidget(
        CardService.newDecoratedText()
          .setTopLabel('Potentiel estimé')
          .setText(String(analysis.montant))
          .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.DOLLAR))
      );
    }

    bizSection.addWidget(
      CardService.newTextParagraph().setText("Résumé :\n" + String(analysis.resume || "Analyse en cours..."))
    );
    card.addSection(bizSection);

    const threadUrl = "https://mail.google.com/mail/u/0/#inbox/" + thread.getId();

    // --- Section Plan d'action ---
    if (analysis.tacheTitre) {
      const taskSection = CardService.newCardSection().setHeader('📋 Action recommandée');

      taskSection.addWidget(
        CardService.newDecoratedText()
          .setText(String(analysis.tacheTitre))
          .setBottomLabel(String(analysis.tacheDescription || "Pas de détail additionnels."))
          .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.CLOCK))
      );

      // Tronquer la description pour éviter le dépassement de limite de paramètres Google
      const taskNotes = `${analysis.tacheDescription || ''}\n\nTiers : ${sender}\nLien : ${threadUrl}`
        .substring(0, 1500);

      taskSection.addWidget(
        CardService.newTextButton()
          .setText('✅ Ajouter à mes tâches')
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(
            CardService.newAction()
              .setFunctionName('createTaskAction')
              .setParameters({ title: String(analysis.tacheTitre), notes: String(taskNotes) })
          )
      );
      card.addSection(taskSection);
    }

    // --- Section Actions rapides ---
    const actionSection = CardService.newCardSection().setHeader('Actions rapides');
    actionSection.addWidget(CardService.newDivider());

    const primaryButtonSet = CardService.newButtonSet()
      .addButton(
        CardService.newTextButton()
          .setText("✉️ Ouvrir l'email")
          .setOpenLink(CardService.newOpenLink().setUrl(threadUrl))
      );

    if (analysis.reponseSuggree) {
      // Stockage en cache pour éviter de dépasser la limite de taille des paramètres
      CacheService.getUserCache().put('draft_' + messageId, String(analysis.reponseSuggree), 1800);

      primaryButtonSet.addButton(
        CardService.newTextButton()
          .setText('📝 Préparer le brouillon')
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(
            CardService.newAction()
              .setFunctionName('createReplyDraftAction')
              .setParameters({ messageId: String(messageId) })
          )
      );
    }
    actionSection.addWidget(primaryButtonSet);

    const secondaryButtonSet = CardService.newButtonSet()
      .addButton(
        CardService.newTextButton()
          .setText('📅 Rappel (Agenda)')
          .setOnClickAction(
            CardService.newAction()
              .setFunctionName('showDatePickerCardAction')
              .setParameters({ subject: String(subject), sender: String(sender) })
          )
      );
    actionSection.addWidget(secondaryButtonSet);
    card.addSection(actionSection);

    card.setFixedFooter(
      CardService.newFixedFooter().setPrimaryButton(
        CardService.newTextButton()
          .setText('Tableau de bord')
          .setOnClickAction(CardService.newAction().setFunctionName('goToHomepageAction'))
      )
    );

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
 * Appelle l'API Gemini pour analyser le contenu d'un email.
 */
function callGeminiAI(text, subject, pdfAttachments = []) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

  if (!apiKey) {
    throw new Error("La clé API GEMINI_API_KEY est introuvable dans les propriétés du script.");
  }

  const modelUrl =
    "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey;

  let parts = [
    {
      text: "Tu es un assistant personnel. Analyse l'email et les PDF joints pour extraire " +
            "les montants, références, intentions et identifier l'action principale à réaliser."
    },
    {
      text: 'Réponds UNIQUEMENT en JSON valide : {"urgence":"...", "sentiment":"Emoji + texte", ' +
            '"categorie":"...", "montant":"...", "numeroCommande":"...", "resume":"...", ' +
            '"reponseSuggree":"...", "tacheTitre":"Titre de l\'action à faire (court)", ' +
            '"tacheDescription":"Détail de l\'action (court)"}'
    },
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
    } catch (e) {
      Logger.log("Erreur PDF: " + e.message);
    }
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

  // Retry avec back-off exponentiel en cas de rate limit (429)
  let response;
  for (let i = 0; i < 3; i++) {
    response = UrlFetchApp.fetch(modelUrl, options);
    if (response.getResponseCode() !== 429) break;
    Utilities.sleep(Math.pow(2, i) * 1000);
  }

  try {
    const resText = response.getContentText();
    const json    = JSON.parse(resText);

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
      urgence:        "Erreur",
      sentiment:      "Erreur",
      categorie:      "Erreur",
      montant:        "N/A",
      resume:         err.message,
      reponseSuggree: ""
    };
  }
}

// ============================================================
// NAVIGATION
// ============================================================

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
  const query        = e.parameters.query    || "is:unread";

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('📁 ' + categoryName));

  const section = CardService.newCardSection();
  const threads  = GmailApp.search(query, 0, 20);

  if (!threads || threads.length === 0) {
    section.addWidget(
      CardService.newTextParagraph().setText("✅ Aucun message en attente dans cette catégorie.")
    );
  } else {
    threads.forEach((t, index) => {
      let rawSender    = t.getMessages()[0].getFrom();
      let cleanSender  = rawSender.split('<')[0].replace(/"/g, '').trim();
      let subject      = String(t.getFirstMessageSubject());
      let displaySubject = subject.length > 45 ? subject.substring(0, 45) + "..." : subject;

      section.addWidget(
        CardService.newDecoratedText()
          .setText(displaySubject)
          .setBottomLabel(cleanSender)
          .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.EMAIL))
          .setOnClickAction(
            CardService.newAction()
              .setFunctionName('viewThreadAnalysisAction')
              .setParameters({ threadId: String(t.getId()) })
          )
      );

      // Séparateur discret entre les éléments, sauf pour le dernier
      if (index < threads.length - 1) {
        section.addWidget(CardService.newDivider());
      }
    });
  }

  card.addSection(section);

  card.setFixedFooter(
    CardService.newFixedFooter().setPrimaryButton(
      CardService.newTextButton()
        .setText('Retour')
        .setOnClickAction(CardService.newAction().setFunctionName('goToHomepageAction'))
    )
  );

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card.build()))
    .build();
}

/**
 * Ouvre la carte d'analyse d'un fil de discussion.
 */
function viewThreadAnalysisAction(e) {
  const msgId = GmailApp.getThreadById(e.parameters.threadId).getMessages()[0].getId();
  return CardService.newActionResponseBuilder()
    .setNavigation(
      CardService.newNavigation().pushCard(
        buildContextualCard({ messageMetadata: { messageId: String(msgId) } })
      )
    )
    .build();
}

// ============================================================
// ACTIONS RAPIDES
// ============================================================

/**
 * Crée un brouillon de réponse dans Gmail.
 */
function createReplyDraftAction(e) {
  try {
    const messageId = e.parameters.messageId;
    
    // 1. On récupère le texte du cache
    const cachedReply = CacheService.getUserCache().get('draft_' + messageId);
    const finalReplyText = cachedReply || "Bonjour,\n\nMerci pour votre message.";

    // 2. On convertit les retours à la ligne textuels en balises HTML
    // Cela garantit que le brouillon Gmail aura de jolis paragraphes
    const htmlFormattedReply = finalReplyText.replace(/\n/g, '<br>');

    // 3. On crée le brouillon en utilisant l'option 'htmlBody'
    GmailApp.getMessageById(messageId).createDraftReply('', {
      htmlBody: htmlFormattedReply
    });
    
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("📝 Brouillon formaté créé !"))
      .build();
      
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("❌ Erreur Brouillon : " + err.message))
      .build();
  }
}

/**
 * Crée une tâche dans Google Tasks avec le lien du mail.
 */
function createTaskAction(e) {
  try {
    const task = {
      title: e.parameters.title,
      notes: e.parameters.notes
    };
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

// ============================================================
// AGENDA
// ============================================================

/**
 * Affiche une sous-carte pour sélectionner la date et l'heure du rappel.
 */
function showDatePickerCardAction(e) {
  const sender  = e.parameters.sender  || "Client";
  const subject = e.parameters.subject || "Sans objet";

  const defaultDate = new Date();
  defaultDate.setDate(defaultDate.getDate() + 1);
  defaultDate.setHours(9, 0, 0, 0);

  const dateTimePicker = CardService.newDateTimePicker()
    .setTitle("Quand voulez-vous être rappelé ?")
    .setFieldName("rappelDateTime")
    .setValueInMsSinceEpoch(defaultDate.getTime());

  const buttonSet = CardService.newButtonSet()
    .addButton(
      CardService.newTextButton()
        .setText("Valider et planifier")
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(
          CardService.newAction()
            .setFunctionName('createCalendarEventAction')
            .setParameters({ subject: subject, sender: sender })
        )
    );

  const section = CardService.newCardSection()
    .addWidget(CardService.newTextParagraph().setText(`Planification pour : **${sender}**`))
    .addWidget(dateTimePicker)
    .addWidget(buttonSet);

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('📅 Planifier un rappel'))
    .addSection(section);

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card.build()))
    .build();
}

/**
 * Création d'un événement dans Google Agenda.
 */
function createCalendarEventAction(e) {
  try {
    const sender  = String(e.parameters.sender  || "Client");
    const subject = String(e.parameters.subject || "Sans objet");

    let startTimeMs = new Date().getTime() + 86400000; // +1 jour par défaut

    if (e.formInput && e.formInput.rappelDateTime) {
      startTimeMs = Number(e.formInput.rappelDateTime.msSinceEpoch || e.formInput.rappelDateTime);
    }

    const startTime = new Date(startTimeMs);
    const endTime   = new Date(startTime.getTime() + 30 * 60000); // +30 minutes

    CalendarApp.getDefaultCalendar().createEvent(
      `Relance : ${sender}`,
      startTime,
      endTime,
      { description: `Sujet : ${subject}` }
    );

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

// ============================================================
// STATISTIQUES & DONNÉES
// ============================================================

/**
 * Récupère les statistiques du tableau de bord (avec cache 1 min).
 */
function getDashboardStats() {
  const cache  = CacheService.getUserCache();
  const cached = cache.get('dashboard_stats');
  if (cached) return JSON.parse(cached);

  const settings = getUserSettings();
  const stats = {
    cat1: GmailApp.search(settings.QUERY_CAT1, 0, 50).length,
    cat2: GmailApp.search(settings.QUERY_CAT2, 0, 50).length,
    cat3: GmailApp.search(settings.QUERY_CAT3, 0, 50).length
  };

  cache.put('dashboard_stats', JSON.stringify(stats), 60);
  return stats;
}

/**
 * Retourne les fils de discussion urgents non lus.
 */
function getUrgentThreads() {
  return GmailApp.search('is:unread is:important', 0, 8).map(t => ({
    threadId: t.getId(),
    subject:  t.getFirstMessageSubject(),
    from:     t.getMessages()[0].getFrom().split('<')[0].replace(/"/g, '').trim()
  }));
}

// ============================================================
// PARAMÈTRES (SETTINGS)
// ============================================================

/**
 * Fonction utilitaire pour générer le menu déroulant d'icônes.
 */
function createIconDropdown(fieldName, title, currentValue) {
  const dropdown = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle(title)
    .setFieldName(fieldName);

  AVAILABLE_ICONS.forEach(icon => {
    dropdown.addItem(icon.label, icon.value, icon.value === currentValue);
  });

  return dropdown;
}

/**
 * Construit la carte des paramètres pour éditer les catégories, requêtes et icônes.
 */
function buildSettingsCard(e) {
  const settings = getUserSettings();

  const section = CardService.newCardSection()
    .addWidget(
      CardService.newTextParagraph().setText(
        "Personnalisez le nom, l'icône et la requête de vos catégories."
      )
    );

  // --- Catégorie 1 ---
  section.addWidget(CardService.newTextInput().setFieldName('LABEL_CAT1').setTitle('Nom de la catégorie 1').setValue(settings.LABEL_CAT1));
  section.addWidget(createIconDropdown('ICON_CAT1', 'Icône 1', settings.ICON_CAT1));
  section.addWidget(CardService.newTextInput().setFieldName('QUERY_CAT1').setTitle('Requête de la catégorie 1').setValue(settings.QUERY_CAT1));
  section.addWidget(CardService.newDivider());

  // --- Catégorie 2 ---
  section.addWidget(CardService.newTextInput().setFieldName('LABEL_CAT2').setTitle('Nom de la catégorie 2').setValue(settings.LABEL_CAT2));
  section.addWidget(createIconDropdown('ICON_CAT2', 'Icône 2', settings.ICON_CAT2));
  section.addWidget(CardService.newTextInput().setFieldName('QUERY_CAT2').setTitle('Requête de la catégorie 2').setValue(settings.QUERY_CAT2));
  section.addWidget(CardService.newDivider());

  // --- Catégorie 3 ---
  section.addWidget(CardService.newTextInput().setFieldName('LABEL_CAT3').setTitle('Nom de la catégorie 3').setValue(settings.LABEL_CAT3));
  section.addWidget(createIconDropdown('ICON_CAT3', 'Icône 3', settings.ICON_CAT3));
  section.addWidget(CardService.newTextInput().setFieldName('QUERY_CAT3').setTitle('Requête de la catégorie 3').setValue(settings.QUERY_CAT3));

  const buttonSet = CardService.newButtonSet()
    .addButton(
      CardService.newTextButton()
        .setText('💾 Enregistrer')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction().setFunctionName('saveSettingsAction'))
    )
    .addButton(
      CardService.newTextButton()
        .setText('🔄 Réinitialiser')
        .setOnClickAction(CardService.newAction().setFunctionName('resetSettingsAction'))
    );

  section.addWidget(buttonSet);

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('⚙️ Paramètres des flux'))
    .addSection(section);

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card.build()))
    .build();
}

/**
 * Sauvegarde les paramètres personnalisés de l'utilisateur.
 */
function saveSettingsAction(e) {
  const formInput = e.formInput;
  const userProps = PropertiesService.getUserProperties();

  if (formInput.LABEL_CAT1) userProps.setProperty('LABEL_CAT1', formInput.LABEL_CAT1);
  if (formInput.LABEL_CAT2) userProps.setProperty('LABEL_CAT2', formInput.LABEL_CAT2);
  if (formInput.LABEL_CAT3) userProps.setProperty('LABEL_CAT3', formInput.LABEL_CAT3);

  if (formInput.QUERY_CAT1) userProps.setProperty('QUERY_CAT1', formInput.QUERY_CAT1);
  if (formInput.QUERY_CAT2) userProps.setProperty('QUERY_CAT2', formInput.QUERY_CAT2);
  if (formInput.QUERY_CAT3) userProps.setProperty('QUERY_CAT3', formInput.QUERY_CAT3);

  if (formInput.ICON_CAT1) userProps.setProperty('ICON_CAT1', formInput.ICON_CAT1);
  if (formInput.ICON_CAT2) userProps.setProperty('ICON_CAT2', formInput.ICON_CAT2);
  if (formInput.ICON_CAT3) userProps.setProperty('ICON_CAT3', formInput.ICON_CAT3);

  CacheService.getUserCache().remove('dashboard_stats');

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText("✅ Paramètres enregistrés !"))
    .setNavigation(CardService.newNavigation().popToRoot().updateCard(buildHomepage(e)))
    .build();
}

/**
 * Remet à zéro tous les paramètres personnalisés.
 */
function resetSettingsAction(e) {
  const userProps = PropertiesService.getUserProperties();

  userProps.deleteProperty('LABEL_CAT1'); userProps.deleteProperty('QUERY_CAT1'); userProps.deleteProperty('ICON_CAT1');
  userProps.deleteProperty('LABEL_CAT2'); userProps.deleteProperty('QUERY_CAT2'); userProps.deleteProperty('ICON_CAT2');
  userProps.deleteProperty('LABEL_CAT3'); userProps.deleteProperty('QUERY_CAT3'); userProps.deleteProperty('ICON_CAT3');

  CacheService.getUserCache().remove('dashboard_stats');

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText("🔄 Paramètres par défaut restaurés"))
    .setNavigation(CardService.newNavigation().popToRoot().updateCard(buildHomepage(e)))
    .build();
}

// ============================================================
// UTILITAIRES
// ============================================================

/**
 * Carte d'erreur générique.
 */
function buildErrorCard(message) {
  const buttonSet = CardService.newButtonSet()
    .addButton(
      CardService.newTextButton()
        .setText("Retour à l'accueil")
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction().setFunctionName('goToHomepageAction'))
    );

  const section = CardService.newCardSection()
    .addWidget(
      CardService.newDecoratedText()
        .setTopLabel("Oups, un problème est survenu")
        .setText(String(message))
        .setWrapText(true)
    )
    .addWidget(buttonSet);

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('❌ Interruption'))
    .addSection(section)
    .build();
}
