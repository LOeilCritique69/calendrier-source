// ============================================================
//   CALENDRIER SORTIES — v41 (Version Optimisée + API JSON)
//   Nouveautés v41 :
//   • API JSON publique (doGet ?format=json) avec filtres
//   • Token d'accès optionnel pour l'API (propriété API_TOKEN)
//   • Invalidation du cache à chaque modification manuelle (onEdit)
//   • Correction de remplirCasesVidesParCroix (statuts vides uniquement)
// ============================================================

function onOpen(e) {
  creerMenuCalendrier();
  mettreEnValeurProchaineSortie();
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() === NOM_FEUILLE_EN_ATTENTE) {
    gererEditionValidation(e, sheet);
  } else if (sheet.getName() === NOM_FEUILLE) {
    // Modification manuelle du calendrier : on vide le cache de l'API
    _invaliderCacheJSON();
  }
}

const MOIS_NOMS = [
  "Janvier","Février","Mars","Avril","Mai","Juin",
  "Juillet","Août","Septembre","Octobre","Novembre","Décembre"
];

const CATEGORIES = ["🎬 Films","📺 Séries","🎮 Jeux","🍿 Animés","🎪 Événements","📦 Autre","🎵 Music"];
const CATEGORIE_LEGACY_FILMS_SERIES = "🎬 Films/Séries";

const PLATEFORMES = [
  "🎥 Cinéma", "📺 Netflix", "🟦 Prime Video", "✨ Disney+", "📡 Canal+",
  "🍎 Apple TV+", "🟣 Max", "🍥 Crunchyroll", "🎌 ADN", "🇯🇵 Wakanim",
  "📻 Molotov", "🌐 Pluto TV", "🎞️ OCS", "⭐ Paramount+", "📺 Twitch", "▶️ YouTube",
  "🎮 PlayStation", "🎮 PS5", "🎮 PS4", "🎮 Xbox", "🎮 Xbox Series", "🎮 Xbox One",
  "🎮 Switch", "🎮 PC", "🎮 PC / Steam", "🌀 Autre streaming", "📦 Non précisé"
];

const YEARS = [2026, 2027, 2028, 2029];
const SAISONS = ["Hiver","Printemps","Été","Automne"];
const SAISON_ORDRE = { "Hiver": 0, "Printemps": 1, "Été": 2, "Automne": 3 };

const MIN_LIGNES_PAR_MOIS = 5;
const MARGE_LIGNES_VIDES  = 3;
const LIGNES_INDETERMINE  = 15;

const OFF_JOUR        = 0;
const OFF_TITRE       = 1;
const OFF_CAT         = 2;
const OFF_PLATEFORME = 3;
const OFF_LIEN        = 4;
const OFF_STATUT      = 5;

const COLS_PAR_MOIS   = 6;
const BLOC_WIDTH      = 7;
const GRID_COLS       = 2;
const GRID_ROWS       = 6;
const GRID_GAP_ROWS   = 1;

const GAP_AVANT_INDET = 1;
const LARGEUR_GRILLE  = GRID_COLS * BLOC_WIDTH - 1;
const HEADER_ROWS     = 5;
const GRID_START_ROW  = HEADER_ROWS + 2;

const PALETTE_MOIS = [
  "#26354A", "#315B73", "#356B63", "#477A68", "#62816A", "#7A7263",
  "#8A645A", "#9A7155", "#685B78", "#765F86", "#41485C", "#8A3F4B"
];

const COULEUR_ANNEE = { 2026: "#182433", 2027: "#243A38", 2028: "#3A2635", 2029: "#332A45" };
const COULEUR_INDETERMINE = "#59636E";

const NOM_FEUILLE                 = "📅 CALENDRIER 2026-2029";
const NOM_FEUILLE_CORBEILLE       = "🗑️ CORBEILLE";
const NOM_FEUILLE_JOURNAL_AMIS   = "🧾 JOURNAL AMIS";
const NOM_FEUILLE_DASHBOARD      = "📋 TABLEAU DE BORD";
const NOM_FEUILLE_EN_ATTENTE     = "⏳ À VALIDER";
const NOM_FEUILLE_WATCHLIST      = "🔭 À SURVEILLER";
const NOM_CALENDRIER_GOOGLE      = "🎬 Sorties — L'Œil Critique";

const PROP_KEY_TMDB               = "TMDB_API_KEY";
const PROP_KEY_IGDB_CLIENT_ID    = "IGDB_CLIENT_ID";
const PROP_KEY_IGDB_CLIENT_SECRET = "IGDB_CLIENT_SECRET";
const PROP_KEY_IGDB_TOKEN        = "IGDB_ACCESS_TOKEN";
const PROP_KEY_IGDB_TOKEN_EXPIRE = "IGDB_TOKEN_EXPIRE";
const PROP_KEY_MEDIASTACK         = "MEDIASTACK_API_KEY";
const PROP_KEY_YOUTUBE           = "YOUTUBE_API_KEY";
const PROP_KEY_GEMINI            = "GEMINI_API_KEY";
const PROP_KEY_GEMINI_MODELE     = "GEMINI_MODEL";
const MODELE_GEMINI_DEFAUT       = "gemini-3.8-flash";
const MODELE_GEMINI_FALLBACK     = "gemini-3.1-flash-lite";
const MODELE_GEMINI_SECOURS      = "gemini-flash-latest";
const PROP_KEY_CODE_AMIS         = "CODE_ACCES_AMIS";
const PROP_KEY_API_TOKEN         = "API_TOKEN";
const PROP_KEY_RECENTS           = "AJOUTS_RECENTS";
const PROP_KEY_HIGHLIGHTS        = "HIGHLIGHTS_PRECEDENTS";
const NB_RECENTS_MAX             = 6;
const MAX_PROPOSITIONS_PAR_HEURE = 30;

const FUSEAU_FR = "Europe/Paris";

const CACHE_KEY_JSON = "calendrier_json_v4";
const CACHE_KEY_PAGE = "calendrier_page_v1";
const CACHE_TTL_SECONDS = 1800;

// ============================================================
//   CACHE & UTILITAIRES
// ============================================================

function _invaliderUnCache(cle) {
  const cache = CacheService.getScriptCache();
  const countStr = cache.get(cle + "_count");
  if (countStr) {
    const count = parseInt(countStr, 10);
    for (let i = 0; i < count; i++) cache.remove(cle + "_" + i);
    cache.remove(cle + "_count");
  }
  cache.remove(cle);
}

function _invaliderCacheJSON() {
  try { _invaliderUnCache(CACHE_KEY_JSON); } catch (e) {}
  try { _invaliderUnCache(CACHE_KEY_PAGE); } catch (e) {}
}

function _sauvegarderCacheDecoupe(jsonStr, cle) {
  cle = cle || CACHE_KEY_JSON;
  try {
    const cache = CacheService.getScriptCache();
    const chunkSize = 20000;
    const chunks = Math.ceil(jsonStr.length / chunkSize);

    cache.put(cle + "_count", String(chunks), CACHE_TTL_SECONDS);
    for (let i = 0; i < chunks; i++) {
      cache.put(cle + "_" + i, jsonStr.substring(i * chunkSize, (i + 1) * chunkSize), CACHE_TTL_SECONDS);
    }
  } catch (e) {
    Logger.log("Erreur lors de la sauvegarde du cache : " + e.toString());
  }
}

function _lireCacheDecoupe(cle) {
  cle = cle || CACHE_KEY_JSON;
  try {
    const cache = CacheService.getScriptCache();
    const countStr = cache.get(cle + "_count");
    if (!countStr) return null;
    const count = parseInt(countStr, 10);
    if (!Number.isInteger(count) || count < 1 || count > 1000) return null;
    let jsonStr = "";
    for (let i = 0; i < count; i++) {
      const chunk = cache.get(cle + "_" + i);
      if (!chunk) return null;
      jsonStr += chunk;
    }
    return jsonStr;
  } catch (e) { return null; }
}

function _normSimple(s) {
  if (!s) return "";
  let t = s.toString().toLowerCase();
  try { t = t.normalize("NFD").replace(/[\u0300-\u036f]/g, ""); } catch(e){}
  return t.trim();
}

function _escapeHtml(str) {
  if (str === null || str === undefined) return "";
  return str.toString().replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
}

function _neutraliserFormule(v) {
  const s = (v === null || v === undefined) ? "" : v.toString();
  return /^[=+\-@\t\r]/.test(s) ? " " + s : s;
}

function _urlSure(u) {
  const s = (u || "").toString().trim();
  return /^https?:\/\/[^\s<>"']+$/i.test(s) ? s.slice(0, 500) : "";
}

// Validation stricte d'une date grégorienne.
function _dateLocaleValide_(annee, mois, jour) {
  if (!Number.isInteger(annee) || !Number.isInteger(mois) || !Number.isInteger(jour)) return false;
  if (annee < 1 || mois < 1 || mois > 12 || jour < 1 || jour > 31) return false;
  const d = new Date(0);
  d.setHours(12, 0, 0, 0);
  d.setFullYear(annee, mois - 1, jour);
  return d.getFullYear() === annee && d.getMonth() === mois - 1 && d.getDate() === jour;
}

function _statutEstVu_(valeur) {
  return ["✔️", "✔", "✓", "✅", "☑️", "☑"].includes((valeur || "").toString().trim());
}

function _exigerContexteTableur() {
  try { SpreadsheetApp.getUi(); }
  catch (e) { throw new Error("Action réservée au propriétaire du tableur."); }
}

function _alerte(msg) {
  try { SpreadsheetApp.getUi().alert(msg); }
  catch (e) { Logger.log(msg); }
}

function _toast(msg, titre, secondes) {
  try { SpreadsheetApp.getActiveSpreadsheet().toast(msg, titre || "Calendrier", secondes || 5); }
  catch (e) { Logger.log(msg); }
}

function _getOrCreateCorbeille(ss) {
  let f = ss.getSheetByName(NOM_FEUILLE_CORBEILLE);
  if (!f) {
    f = ss.insertSheet(NOM_FEUILLE_CORBEILLE);
    f.setTabColor("#e63946");
    f.getRange(1, 1, 1, 7).setValues([["TITRE", "DATE RESTAURATION", "CATÉGORIE", "PLATEFORME", "LIEN", "STATUT", "SUPPRIMÉ LE"]])
      .setFontWeight("bold").setBackground("#2B1D1D").setFontColor("#FFFFFF");
    f.setColumnWidths(1, 7, 140);
    f.setFrozenRows(1);
    f.hideSheet();
  }
  return f;
}

function _getOrCreateWatchlist(ss) {
  let f = ss.getSheetByName(NOM_FEUILLE_WATCHLIST);
  if (!f) {
    f = ss.insertSheet(NOM_FEUILLE_WATCHLIST);
    f.setTabColor("#4f46e5");
    f.getRange(1, 1, 1, 4).setValues([["TITRE À SURVEILLER", "TYPE (film / serie / jeu / anime)", "DERNIÈRE VÉRIFICATION", "DERNIER STATUT"]])
      .setFontWeight("bold").setBackground("#1e1e34").setFontColor("#ffffff");
    f.setColumnWidths(1, 1, 260);
    f.setColumnWidths(2, 1, 220);
    f.setColumnWidths(3, 1, 170);
    f.setColumnWidths(4, 1, 260);
    f.setFrozenRows(1);
    f.getRange(2, 2, 200, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(["film", "serie", "jeu", "anime"], true).setAllowInvalid(true).build()
    );
  }
  return f;
}

// ============================================================
//   INTÉGRATION GEMINI (IA)
// ============================================================

function configurerCleGemini() {
  const ui = SpreadsheetApp.getUi();
  const rep = ui.prompt("🤖 Clé API Gemini", "Colle ta clé API Gemini :", ui.ButtonSet.OK_CANCEL);
  if (rep.getSelectedButton() === ui.Button.OK && rep.getResponseText().trim()) {
    PropertiesService.getScriptProperties().setProperty(PROP_KEY_GEMINI, rep.getResponseText().trim());
    ui.alert("✅ Clé Gemini enregistrée avec succès !");
  }
}

function configurerModeleGemini() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const actuel = props.getProperty(PROP_KEY_GEMINI_MODELE) || MODELE_GEMINI_DEFAUT;
  const rep = ui.prompt(
    "🤖 Modèle Gemini",
    `Modèle actuel : ${actuel}\n\nSaisis le nom d'un modèle (ex : gemini-3.8-flash) ou laisse vide pour revenir au défaut :`,
    ui.ButtonSet.OK_CANCEL
  );
  if (rep.getSelectedButton() !== ui.Button.OK) return;
  const m = rep.getResponseText().trim().replace(/^models\//, "");
  if (m) props.setProperty(PROP_KEY_GEMINI_MODELE, m);
  else props.deleteProperty(PROP_KEY_GEMINI_MODELE);
  ui.alert(`✅ Modèle : ${m || MODELE_GEMINI_DEFAUT}`);
}

function _erreurModeleIndisponible(err) {
  if (!err) return false;
  const msg = (err.message || "").toLowerCase();
  return err.code === 404 || msg.includes("no longer available") || msg.includes("not found") || msg.includes("is not supported");
}

function _appelerGemini(modele, apiKey, payload) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(modele)}:generateContent`;
  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: { "x-goog-api-key": apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  let json;
  try { json = JSON.parse(resp.getContentText()); }
  catch (e) { json = { error: { code: resp.getResponseCode(), message: "Réponse non JSON de l'API Gemini." } }; }
  return json;
}

function analyserTexteAvecGemini(texteBrut) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty(PROP_KEY_GEMINI);
  if (!apiKey) throw new Error("Clé API Gemini non configurée (menu ⚙️ Configuration & API).");

  const aujourdhui = Utilities.formatDate(new Date(), FUSEAU_FR, "dd/MM/yyyy");
  const payload = {
    system_instruction: {
      parts: [{
        text: `Tu extrais les informations d'une annonce de sortie (tweet, post, article) pour un calendrier.
Date du jour : ${aujourdhui}. Réponds UNIQUEMENT par un objet JSON valide, sans texte autour.
Champs :
- "titre" : nom propre de l'œuvre, sans hashtags ni mentions (ex : "Spider-Man: Brand New Day").
- "date" : au format JJ/MM/AAAA si le jour est connu, sinon MM/AAAA, sinon "Saison AAAA" (Hiver, Printemps, Été, Automne), sinon AAAA. Si l'année est absente, déduis la plus probable à partir de la date du jour.
- "categorie" : exactement une valeur parmi ${JSON.stringify(CATEGORIES)}.
- "plateforme" : support ou plateforme, de préférence exactement une valeur parmi ${JSON.stringify(PLATEFORMES)} ; sinon chaîne vide.`
      }]
    },
    contents: [{ parts: [{ text: texteBrut }] }],
    generationConfig: {
      responseMimeType: "application/json",
      responseSchema: {
        type: "OBJECT",
        properties: {
          titre: { type: "STRING" },
          date: { type: "STRING" },
          categorie: { type: "STRING" },
          plateforme: { type: "STRING" }
        },
        required: ["titre", "date", "categorie"]
      },
      temperature: 0.1
    }
  };

  const modelePrincipal = props.getProperty(PROP_KEY_GEMINI_MODELE) || MODELE_GEMINI_DEFAUT;
  const modeles = [...new Set([modelePrincipal, "gemini-1.5-flash", MODELE_GEMINI_SECOURS])];
  const maxTentatives = 4;
  let derniereErreur = null;

  for (const modele of modeles) {
    for (let tentative = 1; tentative <= maxTentatives; tentative++) {
      const json = _appelerGemini(modele, apiKey, payload);

      if (!json.error) {
        const cand = json.candidates && json.candidates[0];
        const texte = cand && cand.content && cand.content.parts && cand.content.parts.map(p => p.text || "").join("");
        if (!texte) {
          const raison = (cand && cand.finishReason) || (json.promptFeedback && json.promptFeedback.blockReason) || "inconnue";
          throw new Error(`Réponse Gemini vide ou bloquée (raison : ${raison}).`);
        }
        return _normaliserExtraction(_lireJsonTolerant(texte));
      }

      derniereErreur = json.error;
      const msg = (json.error.message || "").toLowerCase();
      const surcharge = json.error.code === 429 || json.error.code === 503 || msg.includes("high demand") || msg.includes("quota") || msg.includes("overloaded");

      if (_erreurModeleIndisponible(json.error)) break;

      if (surcharge && tentative < maxTentatives) {
        const jitter = Math.floor(Math.random() * 2000) + 1000;
        const delai = Math.pow(2, tentative) * 1000 + jitter;
        Utilities.sleep(delai);
        continue;
      }
    }
  }

  throw new Error("Surcharge temporaire des serveurs Google (429/503). Réessaie dans quelques secondes !");
}

function _lireJsonTolerant(texte) {
  let t = (texte || "").toString().replace(/```json|```/gi, "").trim();
  try { return JSON.parse(t); }
  catch (e) {
    const debut = t.indexOf("{"), fin = t.lastIndexOf("}");
    if (debut !== -1 && fin > debut) {
      try { return JSON.parse(t.slice(debut, fin + 1)); } catch (e2) {}
    }
    throw new Error("Gemini a renvoyé un JSON illisible (réponse tronquée ?).");
  }
}

function _normaliserExtraction(x) {
  x = x || {};
  const titre = (x.titre || "").toString().replace(/^[#@]+/, "").trim();
  let date = (x.date || "").toString().trim();
  if (!titre) throw new Error("Gemini n'a pas trouvé de titre dans le texte.");
  if (!date) throw new Error("Gemini n'a pas trouvé de date dans le texte.");

  const catBrute = (x.categorie || "").toString().trim();
  const categorie = CATEGORIES.includes(catBrute)
    ? catBrute
    : (CATEGORIES.find(c => _normSimple(c).includes(_normSimple(catBrute).replace(/[^a-z]/g, "")) && catBrute) || CATEGORIES[0]);

  const platBrute = (x.plateforme || "").toString().trim();
  const plateforme = PLATEFORMES.find(p => p === platBrute)
    || PLATEFORMES.find(p => platBrute && _normSimple(p).includes(_normSimple(platBrute).replace(/^[^a-z0-9]+/, "")))
    || platBrute;

  return { titre, date, categorie, plateforme };
}

function ajouterDepuisPostRS() {
  const ui = SpreadsheetApp.getUi();
  const rep = ui.prompt(
    "🤖 Ajouter via annonce Gemini",
    "Colle le texte du tweet / post (ex: 'FIRST LOOK: Spider-Man...'):",
    ui.ButtonSet.OK_CANCEL
  );
  if (rep.getSelectedButton() !== ui.Button.OK) return;
  const texte = rep.getResponseText().trim();
  if (!texte) return;

  try {
    _toast("Gemini analyse l'annonce…", "🤖 IA", 15);
    const x = analyserTexteAvecGemini(texte);

    const conf = ui.alert(
      "🤖 Vérifie avant d'ajouter",
      `• Titre : ${x.titre}\n• Date : ${x.date}\n• Catégorie : ${x.categorie}\n• Support : ${x.plateforme || "—"}\n\nAjouter cette sortie au calendrier ?`,
      ui.ButtonSet.YES_NO
    );
    if (conf !== ui.Button.YES) return;

    const msg = insererEntree_(x.titre, x.date, x.categorie, x.plateforme, "", "❌");
    mettreEnValeurProchaineSortie();
    ui.alert(msg);
  } catch (err) {
    ui.alert("❌ Erreur Gemini : " + err.message);
  }
}

// Fonction générique (non utilisée par le menu) — nécessite que getApiKey_ soit définie si tu t'en sers
function appelerGemini_(prompt, systemInstruction, schema, modele) {
  modele = modele || MODELE_GEMINI_DEFAUT;
  const apiKey = PropertiesService.getScriptProperties().getProperty(PROP_KEY_GEMINI);
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modele}:generateContent?key=${apiKey}`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: "application/json"
    }
  };

  if (systemInstruction) {
    payload.systemInstruction = { parts: [{ text: systemInstruction }] };
  }

  if (schema) {
    payload.generationConfig.responseSchema = schema;
  }

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const status = response.getResponseCode();

  if (status === 429 && modele !== MODELE_GEMINI_FALLBACK) {
    Logger.log(`Quota atteint pour ${modele}. Basculement sur ${MODELE_GEMINI_FALLBACK}...`);
    return appelerGemini_(prompt, systemInstruction, schema, MODELE_GEMINI_FALLBACK);
  }

  if (status !== 200) {
    throw new Error(`Erreur API Gemini (${status}): ${response.getContentText()}`);
  }

  return JSON.parse(response.getContentText());
}

// ============================================================
//   ANALYSE & NORMALISATION DE TITRES
// ============================================================

function _extraireNomBaseAnime(titreBrut, eclater) {
  let t = (titreBrut || "").toString().trim();
  t = t.replace(/\[.*?\]/g, " ");
  t = t.replace(/\s*[-–—]\s*(ep|ép|jour|part|cour).*$/i, " ");
  t = t.replace(/\b(saison|season|cour|partie|part)\b.*$/i, " ");
  t = t.replace(/\b(ep|ép|episode|épisode)\.?\s*\d+.*$/i, " ");
  t = t.replace(/\bs\d+\b.*$/i, " ");
  if (eclater !== false) t = t.replace(/([a-z])([A-Z])/g, "$1 $2");
  t = t.replace(/[:\-–—\s]+$/, "");
  return t.replace(/\s+/g, " ").trim();
}

function _essaisRechercheAnime(titreOriginal) {
  const base = _extraireNomBaseAnime(titreOriginal);
  if (!base) return [titreOriginal];
  const brut = _extraireNomBaseAnime(titreOriginal, false);
  const compact = brut.replace(/[\s:]/g, "").toLowerCase();
  const essais = [base];
  if (compact === "rezero") essais.unshift("Re:Zero");
  else if (compact === "snk") essais.unshift("Attack on Titan", "Shingeki no Kyojin");
  else if (compact === "danmachi") essais.unshift("DanMachi");
  if (brut && brut !== base) essais.push(brut);
  return [...new Set(essais.filter(Boolean))];
}

function _normaliserTitre(s) {
  if (!s) return "";
  let t = s.toString().toLowerCase();
  try { t = t.normalize("NFD").replace(/[\u0300-\u036f]/g, ""); } catch (e) {}
  t = t.replace(/\[.*?\]/g, " ");
  t = t.replace(/\s*-\s*ep(isode)?\.?\s*\d+/gi, " ");
  t = t.replace(/\bep(isode)?\.?\s*\d+\b/gi, " ");
  t = t.replace(/\s*-\s*jour\s*\d+/gi, " ");
  t = t.replace(/\b(cour|partie|part|pt)\.?\s*\d+\b/gi, " ");
  t = t.replace(/\b(vostfr|vost|vf|vff|vfq|uncut|oav|ova|ona)\b/gi, " ");
  t = t.replace(/[:\/]/g, " ");
  t = t.replace(/\s+[-–—]\s+/g, " ");
  t = t.replace(/[^a-z0-9\s\-]/g, " ");
  return t.replace(/\s+/g, " ").trim();
}

function _essaisRecherche(titreOriginal) {
  const t = (titreOriginal || "").toString().trim();
  const norm = _normaliserTitre(t);
  const mSaison = t.match(/\b(?:saison|season|s)\s*(\d+)\b/i);
  const numSaison = mSaison ? mSaison[1] : null;
  const titreRacine = norm.replace(/\b(saison|season|s)\s*\d+\b/gi, "").replace(/\s+/g, " ").trim();

  const essais = [];
  if (numSaison && titreRacine) {
    essais.push(`${titreRacine} Season ${numSaison}`);
    essais.push(`${titreRacine} ${numSaison}`);
  }
  if (norm) essais.push(norm);
  if (titreRacine && titreRacine !== norm) essais.push(titreRacine);
  return [...new Set(essais.filter(Boolean))];
}

function _extraireTitreSansEpisode(titre) {
  let t = (titre || "").toString().trim();
  t = t.replace(/\s*[-–—:]\s*(ep|ép|episode|épisode|jour|partie|part|pt)\.?\s*\d+.*$/i, "");
  t = t.replace(/\s+(ep|ép|episode|épisode)\.?\s*\d+\s*$/i, "");
  return t.trim();
}

// ============================================================
//   MOTEUR DYNAMIQUE (Grille & Repérage Optimisé)
// ============================================================

function _infoMois(sheetOrData, annee, m) {
  const fullData = (sheetOrData instanceof Array) ? sheetOrData : sheetOrData.getDataRange().getValues();
  const gridCol = m % GRID_COLS;
  const colDebut = 1 + gridCol * BLOC_WIDTH;
  const targetNorm = _normSimple(MOIS_NOMS[m]);

  let rAnnee = -1;
  for (let r = 0; r < fullData.length; r++) {
    const val = _normSimple(fullData[r][0]);
    if (val.includes("sorties") && val.includes(String(annee))) { rAnnee = r; break; }
  }
  if (rAnnee === -1) return { ligneTitreMois: -1, ligneHeaders: -1, ligneData: -1, nbLignesData: MIN_LIGNES_PAR_MOIS, colDebut };

  const cIdx = colDebut - 1;
  let rMois = -1;
  for (let r = rAnnee; r < fullData.length; r++) {
    const valGauche = _normSimple(fullData[r][0]);
    if (r > rAnnee && valGauche.includes("sorties") && !valGauche.includes(String(annee))) break;
    if (_normSimple(fullData[r][cIdx]) === targetNorm) { rMois = r; break; }
  }
  if (rMois === -1) return { ligneTitreMois: -1, ligneHeaders: -1, ligneData: -1, nbLignesData: MIN_LIGNES_PAR_MOIS, colDebut };

  const ligneTitreMois = rMois + 1;
  const ligneData = ligneTitreMois + 2;

  let rNext = -1;
  for (let r = rMois + 2; r < fullData.length; r++) {
    const txtLeftNorm = _normSimple(fullData[r][0]);
    const txtRightNorm = _normSimple(fullData[r][BLOC_WIDTH] || "");
    const estEntete = MOIS_NOMS.some(nm => _normSimple(nm) === txtLeftNorm || _normSimple(nm) === txtRightNorm) ||
                      txtLeftNorm.includes("date indeterminee") ||
                      txtLeftNorm.includes("sorties");
    if (estEntete) { rNext = r; break; }
  }

  let nbLignesData = MIN_LIGNES_PAR_MOIS;
  if (rNext !== -1) {
    const rowBefore = fullData[rNext - 1] || [];
    if (rowBefore.slice(0, LARGEUR_GRILLE).every(v => v === "")) {
      nbLignesData = (rNext - 1) - (rMois + 2);
    } else {
      nbLignesData = rNext - (rMois + 2);
    }
  } else {
    let lastRow = fullData.length - 1;
    while (lastRow >= rMois + 2) {
      if (!fullData[lastRow].slice(0, LARGEUR_GRILLE).every(v => v === "")) break;
      lastRow--;
    }
    nbLignesData = Math.max(MIN_LIGNES_PAR_MOIS, lastRow - (rMois + 2) + 1);
  }

  return { ligneTitreMois, ligneHeaders: ligneTitreMois + 1, ligneData, nbLignesData: Math.max(1, nbLignesData), colDebut };
}

function _infoIndetermine(sheetOrData, annee) {
  const fullData = (sheetOrData instanceof Array) ? sheetOrData : sheetOrData.getDataRange().getValues();

  let rAnnee = -1;
  for (let r = 0; r < fullData.length; r++) {
    const val = _normSimple(fullData[r][0]);
    if (val.includes("sorties") && val.includes(String(annee))) { rAnnee = r; break; }
  }

  let rIndet = -1;
  if (rAnnee !== -1) {
    for (let r = rAnnee; r < fullData.length; r++) {
      if (_normSimple(fullData[r][0]).includes("date indeterminee")) { rIndet = r; break; }
    }
  }
  if (rIndet === -1) return { ligneTitre: -1, ligneHeaders: -1, ligneData: -1, nbLignesData: LIGNES_INDETERMINE, colDebut: 1 };

  const ligneTitre = rIndet + 1;
  const ligneData = ligneTitre + 2;

  let rNext = -1;
  for (let r = rIndet + 2; r < fullData.length; r++) {
    const val = _normSimple(fullData[r][0]);
    if (val.includes("sorties")) { rNext = r; break; }
  }

  let lastNonEmpty = (rNext !== -1 ? rNext - 1 : fullData.length - 1);
  while (lastNonEmpty >= ligneData && fullData[lastNonEmpty].slice(0, COLS_PAR_MOIS).every(v => v === "")) {
    lastNonEmpty--;
  }

  const nbLignesData = Math.max(LIGNES_INDETERMINE, lastNonEmpty - ligneData + 1);
  return { ligneTitre, ligneHeaders: ligneTitre + 1, ligneData, nbLignesData, colDebut: 1 };
}

function _ligneVide(couleur) { return Array(COLS_PAR_MOIS).fill(couleur); }

function _parseDateEntree(dateStr) {
  if (!dateStr) return { type: "invalide", raison: "Date vide" };
  if (dateStr instanceof Date && !isNaN(dateStr.getTime())) {
    const annee = dateStr.getFullYear();
    if (!YEARS.includes(annee)) return { type: "invalide", raison: `Année ${annee} hors calendrier (${YEARS.join(", ")})` };
    return { type: "mois", jour: String(dateStr.getDate()).padStart(2, "0"), mois: dateStr.getMonth() + 1, annee };
  }

  let s = dateStr.toString().trim();

  const mIso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/);
  if (mIso) {
    const aa = parseInt(mIso[1], 10), mm = parseInt(mIso[2], 10), jj = parseInt(mIso[3], 10);
    if (mm < 1 || mm > 12) return { type: "invalide", raison: "Mois invalide" };
    if (jj < 1 || jj > 31) return { type: "invalide", raison: "Jour invalide" };
    if (!YEARS.includes(aa)) return { type: "invalide", raison: `Année ${aa} hors calendrier (${YEARS.join(", ")})` };
    if (!_dateLocaleValide_(aa, mm, jj)) return { type: "invalide", raison: "Date inexistante : vérifie le jour et le mois" };
    return { type: "mois", jour: String(jj).padStart(2, "0"), mois: mm, annee: aa };
  }

  const dParsed = new Date(s);
  if (s.length > 15 && !isNaN(dParsed.getTime()) && /GMT|UTC|[A-Z][a-z]{2}\s+[A-Z][a-z]{2}/.test(s)) {
    const annee = dParsed.getFullYear();
    if (!YEARS.includes(annee)) return { type: "invalide", raison: `Année ${annee} hors calendrier (${YEARS.join(", ")})` };
    return { type: "mois", jour: String(dParsed.getDate()).padStart(2, "0"), mois: dParsed.getMonth() + 1, annee };
  }

  for (const saison of SAISONS) {
    const m = s.match(new RegExp("^" + saison + "\\s+(\\d{4})$", "i"));
    if (m) {
      const annee = parseInt(m[1], 10);
      if (!YEARS.includes(annee)) return { type: "invalide", raison: `Année ${annee} non prise en compte` };
      return { type: "saison", saison, annee };
    }
  }

  const sNorm = _normSimple(s);
  for (let i = 0; i < MOIS_NOMS.length; i++) {
    const mTextSeul = sNorm.match(new RegExp("^" + _normSimple(MOIS_NOMS[i]) + "\\s+(\\d{4})$"));
    if (mTextSeul) {
      const annee = parseInt(mTextSeul[1], 10);
      if (!YEARS.includes(annee)) return { type: "invalide", raison: `Année ${annee} non prise en charge` };
      return { type: "mois", jour: "?", mois: i + 1, annee };
    }
    const mJourText = sNorm.match(new RegExp("^(\\d{1,2})\\s+" + _normSimple(MOIS_NOMS[i]) + "\\s+(\\d{4})$"));
    if (mJourText) {
      const j = parseInt(mJourText[1], 10);
      const annee = parseInt(mJourText[2], 10);
      if (j < 1 || j > 31) return { type: "invalide", raison: "Jour invalide" };
      if (!YEARS.includes(annee)) return { type: "invalide", raison: `Année ${annee} non prise en charge` };
      if (!_dateLocaleValide_(annee, i + 1, j)) return { type: "invalide", raison: "Date inexistante : vérifie le jour et le mois" };
      return { type: "mois", jour: String(j).padStart(2, "0"), mois: i + 1, annee };
    }
  }

  if (/^\d{4}$/.test(s)) {
    const annee = parseInt(s, 10);
    if (!YEARS.includes(annee)) return { type: "invalide", raison: `Année ${annee} hors limites` };
    return { type: "annee", annee };
  }

  const mQ = s.match(/^\?[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (mQ) {
    const mm = parseInt(mQ[1], 10), aa = parseInt(mQ[2], 10);
    if (mm < 1 || mm > 12) return { type: "invalide", raison: "Mois invalide" };
    if (!YEARS.includes(aa)) return { type: "invalide", raison: `Année ${aa} hors calendrier` };
    return { type: "mois", jour: "?", mois: mm, annee: aa };
  }

  const parts = s.split(/[\/\-.]/);
  if (parts.length === 2) {
    const m = parseInt(parts[0], 10);
    const a = parseInt(parts[1], 10);
    if (m >= 1 && m <= 12 && YEARS.includes(a)) return { type: "mois", jour: "?", mois: m, annee: a };
  }
  if (parts.length === 3) {
    const j = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10);
    const a = parseInt(parts[2], 10);
    if (isNaN(j) || j < 1 || j > 31) return { type: "invalide", raison: "Jour invalide" };
    if (isNaN(m) || m < 1 || m > 12) return { type: "invalide", raison: "Mois invalide" };
    if (!YEARS.includes(a)) return { type: "invalide", raison: `Année ${a} hors calendrier` };
    if (!_dateLocaleValide_(a, m, j)) return { type: "invalide", raison: "Date inexistante : vérifie le jour et le mois" };
    return { type: "mois", jour: String(j).padStart(2, "0"), mois: m, annee: a };
  }

  return { type: "invalide", raison: "Format non reconnu (utiliser JJ/MM/AAAA, MM/AAAA ou Saison AAAA)" };
}

function _cleSortPeriode(periodeStr) {
  const s = (periodeStr || "").toString().trim();
  for (const saison of SAISONS) {
    if (s.indexOf(saison) === 0) return SAISON_ORDRE[saison];
  }
  return 99;
}

function _couleurMois(i) { return PALETTE_MOIS[i % 12]; }
function _assombrir(hex) { return "#" + [0,2,4].map(k => Math.round(parseInt(hex.slice(k+1, k+3), 16) * 0.62).toString(16).padStart(2,"0")).join(""); }
function _bgCouleurCat(cat) {
  return {
    "🎬 Films": "#E8EEF8", "📺 Séries": "#E5F2F1", "🎮 Jeux": "#E8F0E8",
    "🍿 Animés": "#FFF1DF", "🎪 Événements": "#F5E7EC", "📦 Autre": "#ECEDEF", "🎵 Music": "#EEE8F6"
  }[cat] || "#FFFFFF";
}

function _appliquerValidationCategorie(range) { range.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(CATEGORIES, true).setAllowInvalid(true).build()); }
function _appliquerValidationPlateforme(range) { range.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(PLATEFORMES, true).setAllowInvalid(true).build()); }

function _estFilmOuSerie(cat) { return cat === "🎬 Films" || cat === "📺 Séries" || cat === CATEGORIE_LEGACY_FILMS_SERIES; }
function _estJeu(cat) { return cat === "🎮 Jeux"; }
function _estAnime(cat) {
  const c = (cat || "").toString().toLowerCase();
  return c.includes("anime") || c.includes("animé") || c.includes("🍿");
}

function _mapProviderVersPlateforme(nomProvider) {
  const n = (nomProvider || "").toLowerCase();
  if (n.includes("netflix")) return "📺 Netflix";
  if (n.includes("prime video") || n.includes("amazon")) return "🟦 Prime Video";
  if (n.includes("disney")) return "✨ Disney+";
  if (n.includes("canal+")) return "📡 Canal+";
  if (n.includes("apple tv")) return "🍎 Apple TV+";
  if (n.includes("max") || n.includes("hbo")) return "🟣 Max";
  if (n.includes("crunchyroll")) return "🍥 Crunchyroll";
  if (n === "adn" || n.includes("animation digital network")) return "🎌 ADN";
  if (n.includes("wakanim")) return "🇯🇵 Wakanim";
  return nomProvider ? `🌀 ${nomProvider}` : "🌀 Autre streaming";
}

function _mapPlateformesJeu(liste) {
  if (!liste || liste.length === 0) return "";
  const courtes = liste.slice(0, 3).map(p => {
    const n = (p || "").toLowerCase();
    if (n.includes("playstation 5")) return "PS5";
    if (n.includes("playstation 4")) return "PS4";
    if (n.includes("xbox series")) return "Xbox Series";
    if (n.includes("nintendo switch")) return "Switch";
    if (n.includes("pc")) return "PC";
    return p;
  });
  return "🎮 " + [...new Set(courtes)].join(" / ");
}

// ============================================================
//   PARCOURS GLOBAL (Mémoire Optimisé)
// ============================================================

function _parcourirEntrees(sheetOrData, matchFn) {
  const resultats = [];
  const fullData = (sheetOrData instanceof Array) ? sheetOrData : sheetOrData.getDataRange().getValues();

  YEARS.forEach(annee => {
    for (let m = 0; m < 12; m++) {
      const info = _infoMois(fullData, annee, m);
      if (info.ligneData === -1) continue;
      const rStart = info.ligneData - 1;
      const cStart = info.colDebut - 1;

      for (let r = 0; r < info.nbLignesData; r++) {
        if (!fullData[rStart + r]) continue;
        const titre = (fullData[rStart + r][cStart + OFF_TITRE] || "").toString().trim();
        if (!titre) continue;

        const entry = {
          titre, annee, mois: m, indetermine: false,
          jourActuel: (fullData[rStart + r][cStart + OFF_JOUR] || "").toString().trim(),
          cat: (fullData[rStart + r][cStart + OFF_CAT] || "").toString().trim(),
          plateforme: (fullData[rStart + r][cStart + OFF_PLATEFORME] || "").toString().trim(),
          lien: (fullData[rStart + r][cStart + OFF_LIEN] || "").toString().trim(),
          statut: (fullData[rStart + r][cStart + OFF_STATUT] || "").toString().trim(),
          ligneAbs: info.ligneData + r, colDebut: info.colDebut
        };
        if (!matchFn || matchFn(entry)) resultats.push(entry);
      }
    }

    const infoI = _infoIndetermine(fullData, annee);
    if (infoI.ligneData !== -1) {
      const rStartI = infoI.ligneData - 1;
      for (let r = 0; r < infoI.nbLignesData; r++) {
        if (!fullData[rStartI + r]) continue;
        const titre = (fullData[rStartI + r][OFF_TITRE] || "").toString().trim();
        if (!titre) continue;

        const entry = {
          titre, annee, mois: null, indetermine: true,
          periodeActuelle: (fullData[rStartI + r][OFF_JOUR] || "").toString().trim(),
          cat: (fullData[rStartI + r][OFF_CAT] || "").toString().trim(),
          plateforme: (fullData[rStartI + r][OFF_PLATEFORME] || "").toString().trim(),
          lien: (fullData[rStartI + r][OFF_LIEN] || "").toString().trim(),
          statut: (fullData[rStartI + r][OFF_STATUT] || "").toString().trim(),
          ligneAbs: infoI.ligneData + r, colDebut: 1
        };
        if (!matchFn || matchFn(entry)) resultats.push(entry);
      }
    }
  });

  return resultats;
}

// ============================================================
//   MENU CALENDRIER
// ============================================================

function creerMenuCalendrier(ui) {
  if (!ui) ui = SpreadsheetApp.getUi();

  const menuConfig = ui.createMenu("⚙️ Configuration & API")
    .addItem("🌐 Configurer l'URL du Calendrier Web", "configurerUrlWebApp")
    .addItem("🔐 Configurer le token de l'API JSON", "configurerTokenAPI")
    .addItem("🧪 Tester l'API JSON", "testerAPI")
    .addSeparator()
    .addItem("🤖 Configurer la clé Gemini (IA)", "configurerCleGemini")
    .addItem("🤖 Choisir le modèle Gemini", "configurerModeleGemini")
    .addItem("🔑 Configurer la clé TMDB", "configurerCleTMDB")
    .addItem("🔑 Configurer la clé IGDB", "configurerCleIGDB")
    .addItem("📰 Configurer la clé Mediastack (News)", "configurerCleMediastack")
    .addItem("🎥 Configurer la clé YouTube", "configurerCleYoutube")
    .addItem("🔑 Configurer le code d'accès amis", "configurerCodeAmis")
    .addSeparator()
    .addItem("⏰ Activer la MAJ quotidienne", "activerTriggerQuotidien")
    .addItem("🛑 Désactiver la MAJ quotidienne", "desactiverTriggerQuotidien")
    .addItem("⏰ MAJ quotidienne COMPLÈTE (surbrillance + agenda)", "activerMajQuotidienneComplete")
    .addItem("✉️ Activer les rappels email", "activerRappelEmail")
    .addItem("✉️ Désactiver les rappels email", "desactiverRappelEmail")
    .addSeparator()
    .addItem("🤖 Activer la recherche auto hebdo (Watchlist)", "activerSyncAutoHebdo")
    .addItem("🛑 Désactiver la recherche auto hebdo", "desactiverSyncAutoHebdo");

  const menuActions = ui.createMenu("🛠️ Actions & Outils")
    .addItem("⚡ Générateur d'épisodes avancé", "ouvrirGenerateurEpisodes")
    .addItem("⏩ Décaler une série (retard / avance)", "decalerEpisodesSerie")
    .addItem("🎨 Griser les sorties vues (✔️)", "appliquerMiseEnFormeConditionnelle")
    .addItem("⏳ Voir / Valider les propositions d'amis", "voirFeuilleEnAttente")
    .addItem("❌ Remplir les statuts vides avec ❌", "remplirCasesVidesParCroix")
    .addItem("🔄 Retrier tous les mois", "retrierTousLesMois")
    .addItem("✏️ Modifier une date", "modifierDate")
    .addItem("✅ Modifier le statut (Vu / Pas vu)", "marquerStatutSortie")
    .addItem("🗑️ Supprimer une sortie", "supprimerSortie")
    .addItem("♻️ Restaurer une sortie", "restaurerSortie")
    .addItem("🧾 Voir le journal des ajouts (amis)", "voirJournalAmis")
    .addItem("📏 Redimensionner la grille", "ajusterLargeurColonnes")
    .addItem("📤 Exporter en CSV (Drive)", "exporterCalendrierVersCSV");

  ui.createMenu("🗓️ Calendrier")
    .addItem("➕ Ajouter une sortie", "ouvrirFormulaire")
    .addItem("🤖 Ajouter depuis un post / tweet (IA)", "ajouterDepuisPostRS")
    .addItem("🗓️ Vue Agenda (chronologique)", "genererVueAgenda")
    .addItem("🔥 Voir les sorties du jour", "allerAuxSortiesDuJour")
    .addItem("🎯 Aller au mois en cours", "allerAuMoisEnCours")
    .addItem("🔎 Rechercher une sortie", "rechercherSortieRapide")
    .addItem("🧭 Ouvrir l'Explorateur", "ouvrirExplorateur")
    .addSeparator()
    .addItem("🚀 ENRICHIR (Animés, Films, Jeux ou TOUT)", "enrichirAuChoix")
    .addItem("🍿 Enrichir Animés (AniList)", "recupererInfosAnimesAniList")
    .addItem("🖼️ Enrichir Films/Séries (TMDB)", "recupererInfosTMDB")
    .addItem("🎮 Enrichir Jeux (IGDB)", "recupererInfosJeuxIGDB")
    .addItem("📰 Lier des articles aux sorties (Mediastack)", "enrichirLiensAvecMediastack")
    .addItem("🎥 Lier des bandes-annonces (YouTube)", "enrichirLiensAvecYoutube")
    .addItem("🗞️ Générer l'onglet Actualités (Mediastack)", "genererOngletNews")
    .addSeparator()
    .addItem("🔭 Ajouter un titre à surveiller (Watchlist)", "ajouterALaWatchlist")
    .addItem("📋 Voir ma Watchlist", "voirWatchlist")
    .addItem("🤖 Rechercher les nouvelles sorties (Watchlist → Calendrier)", "synchroniserSortiesAutomatiquement")
    .addSeparator()
    .addItem("📋 Actualiser le Tableau de bord", "actualiserTableauDeBord")
    .addItem("📆 Synchroniser Google Calendar", "synchroniserGoogleCalendar")
    .addItem("🌐 Lien du calendrier public (Web App)", "obtenirLienCalendrierAmis")
    .addItem("📱 Lien page mobile", "obtenirLienPageMobile")
    .addSeparator()
    .addSubMenu(menuActions)
    .addSubMenu(menuConfig)
    .addSeparator()
    .addItem("✨ Initialiser le calendrier", "initialiserCalendrier")
    .addToUi();
}

function ouvrirFormulaire() {
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile('Formulaire').setTitle('Ajouter une sortie').setWidth(520));
}

function allerAuxSortiesDuJour() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOM_FEUILLE);
  if (!sheet) return;

  const d = new Date();
  const annee = d.getFullYear();
  const m = d.getMonth();
  const jour = d.getDate();

  const fullData = sheet.getDataRange().getValues();
  const info = _infoMois(fullData, annee, m);
  if (info.ligneData === -1) return SpreadsheetApp.getUi().alert("ℹ️ Mois actuel non trouvé dans le calendrier.");

  const cIdx = info.colDebut - 1;
  const sortiesAujourdhui = [];

  for (let r = 0; r < info.nbLignesData; r++) {
    const row = fullData[info.ligneData - 1 + r];
    if (!row) continue;
    const t = (row[cIdx + OFF_TITRE] || "").toString().trim();
    const j = parseInt((row[cIdx + OFF_JOUR] || "").toString().trim(), 10);
    if (t && j === jour) {
      sortiesAujourdhui.push({ titre: t, ligne: info.ligneData + r, col: info.colDebut });
    }
  }

  if (sortiesAujourdhui.length > 0) {
    sheet.setActiveRange(sheet.getRange(sortiesAujourdhui[0].ligne, sortiesAujourdhui[0].col));
    SpreadsheetApp.getUi().alert(`🎉 ${sortiesAujourdhui.length} sortie(s) aujourd'hui :\n\n` + sortiesAujourdhui.map(s => `• ${s.titre}`).join("\n"));
  } else {
    SpreadsheetApp.getUi().alert("ℹ️ Aucune sortie aujourd'hui. Redirection vers le mois en cours.");
    allerAuMoisEnCours();
  }
}

function allerAuMoisEnCours() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOM_FEUILLE);
  if (!sheet) return;
  const d = new Date();
  const annee = d.getFullYear();
  const m = d.getMonth();
  if (!YEARS.includes(annee)) return SpreadsheetApp.getUi().alert(`❌ L'année en cours (${annee}) n'est pas dans le calendrier.`);

  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  const info = _infoMois(sheet, annee, m);
  if (info.ligneTitreMois !== -1) {
    sheet.getRange(info.ligneTitreMois, info.colDebut).activate();
  }
}

function ajusterLargeurColonnes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOM_FEUILLE);
  if (!sheet) return;
  for (let c = 0; c < GRID_COLS; c++) {
    const colDebut = 1 + c * BLOC_WIDTH;
    sheet.setColumnWidth(colDebut + OFF_JOUR,        50);
    sheet.setColumnWidth(colDebut + OFF_TITRE,      270);
    sheet.setColumnWidth(colDebut + OFF_CAT,        125);
    sheet.setColumnWidth(colDebut + OFF_PLATEFORME, 150);
    sheet.setColumnWidth(colDebut + OFF_LIEN,        160);
    sheet.setColumnWidth(colDebut + OFF_STATUT,      45);
    if (c < GRID_COLS - 1) sheet.setColumnWidth(colDebut + COLS_PAR_MOIS, 18);
  }
  _alerte("✅ Largeurs réappliquées !");
}

// ============================================================
//   INITIALISATION & NAVIGATION FIGÉE
// ============================================================

function initialiserCalendrier() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  let feuilleARemplacer = ss.getSheetByName(NOM_FEUILLE);

  if (feuilleARemplacer) {
    const confirmation = ui.alert(
      "⚠️ Réinitialiser le calendrier ?",
      "Le calendrier va être entièrement reconstruit. Toutes tes données seront sauvegardées et réinjectées proprement. Continuer ?",
      ui.ButtonSet.YES_NO
    );
    if (confirmation !== ui.Button.YES) return;
  }

  let sauvegarde = [];
  const horodatage = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH'h'mm");

  if (feuilleARemplacer) {
    sauvegarde = _sauvegarderDonneesGenerique(feuilleARemplacer);
    feuilleARemplacer.setName(`🗄️ BACKUP_${horodatage}`);
    feuilleARemplacer.hideSheet();
  } else {
    YEARS.forEach(annee => {
      let s = ss.getSheetByName("🗓️ " + annee);
      if (s) {
        sauvegarde.push(..._sauvegarderDonneesGenerique(s));
        s.setName(`🗄️ BACKUP_${horodatage}_${annee}`);
        s.hideSheet();
      }
    });
  }

  const compteurs = {};
  sauvegarde.forEach(d => {
    const p = _parseDateEntree(d.date);
    if (p.type === "mois") {
      const k = `${p.annee}_${p.mois - 1}`;
      compteurs[k] = (compteurs[k] || 0) + 1;
    } else if (p.type === "saison" || p.type === "annee") {
      const k = `${p.annee}_indet`;
      compteurs[k] = (compteurs[k] || 0) + 1;
    }
  });

  const fCal = ss.insertSheet(NOM_FEUILLE);
  fCal.setTabColor("#e63946");
  fCal.setHiddenGridlines(true);

  const colsNecessaires = LARGEUR_GRILLE + 1;
  if (fCal.getMaxColumns() < colsNecessaires) {
    fCal.insertColumnsAfter(fCal.getMaxColumns(), colsNecessaires - fCal.getMaxColumns());
  }

  ajusterLargeurColonnes();
  _creerHeader(fCal);

  let rowCurseur = GRID_START_ROW;
  let blocInfoMap = {};
  let yearRows = {};

  YEARS.forEach(annee => {
    yearRows[annee] = rowCurseur;
    rowCurseur = _creerBlocAnneeDynamique(fCal, annee, rowCurseur, compteurs, blocInfoMap);
  });

  _ajouterNavBarAnnees(fCal, yearRows);

  const res = _injecterSauvegardeEnBloc(fCal, sauvegarde, blocInfoMap);

  _getOrCreateFeuilleEnAttente(ss);
  _getOrCreateCorbeille(ss);
  _getOrCreateWatchlist(ss);
  appliquerMiseEnFormeConditionnelle(true);
  mettreEnValeurProchaineSortie();
  SpreadsheetApp.flush();

  ui.alert(`✅ Calendrier reconstruit sur une feuille unique !\n${res.total - res.echecs} sorties réinjectées.` + (res.echecs > 0 ? `\n\n⚠️ ${res.echecs} doublons / dates invalides ignorés (détail dans le BACKUP).` : ""));
}

function _creerHeader(sheet) {
  sheet.setRowHeight(1, 50);
  sheet.getRange(1, 1, 1, LARGEUR_GRILLE).merge()
    .setValue("📅  CALENDRIER DES SORTIES")
    .setBackground("#182433").setFontColor("#FFFFFF")
    .setFontSize(22).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  sheet.setRowHeight(2, 30);
  sheet.getRange(2, 1, 1, LARGEUR_GRILLE).merge()
    .setValue(`${YEARS.join(" • ")}   —   ${CATEGORIES.join("    ")}   —   ➕ Menu : 🗓️ Calendrier`)
    .setBackground("#F1F3F5").setFontColor("#68727D")
    .setFontSize(10).setFontStyle("italic")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  sheet.setRowHeight(3, 10);
  sheet.getRange(3, 1, 1, LARGEUR_GRILLE).setBackground("#FFFFFF");

  sheet.setRowHeight(4, 34);
  sheet.setRowHeight(5, 16);
  sheet.getRange(4, 1, 2, LARGEUR_GRILLE).merge()
    .setValue("⭐ Prochaine sortie : —")
    .setBackground("#FFF3CD").setFontColor("#765600")
    .setFontWeight("bold").setFontSize(14)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
}

function _ajouterNavBarAnnees(sheet, yearRows) {
  const sheetId = sheet.getSheetId();
  let texteFinal = "";
  let ranges = [];

  YEARS.forEach((annee, i) => {
    let label = "🗓️ ALLER À L'ANNÉE " + annee;
    let start = texteFinal.length;
    texteFinal += label;
    let end = texteFinal.length;
    ranges.push({ start, end, url: `#gid=${sheetId}&range=A${yearRows[annee]}` });
    if (i < YEARS.length - 1) texteFinal += "    ▶    ";
  });

  const rtv = SpreadsheetApp.newRichTextValue().setText(texteFinal);
  ranges.forEach(r => {
    rtv.setLinkUrl(r.start, r.end, r.url);
    rtv.setTextStyle(r.start, r.end, SpreadsheetApp.newTextStyle().setForegroundColor("#F5B82E").setBold(true).build());
  });

  sheet.setRowHeight(3, 30);
  sheet.getRange(3, 1, 1, LARGEUR_GRILLE).merge()
    .setRichTextValue(rtv.build())
    .setBackground("#1a1a2e").setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  sheet.setFrozenRows(5);
}

function _creerBlocAnneeDynamique(sheet, annee, rowStart, compteurs, blocInfoMap) {
  sheet.setRowHeight(rowStart, 48);
  sheet.getRange(rowStart, 1, 1, LARGEUR_GRILLE).merge()
    .setValue(`🗓️  SORTIES ${annee}`)
    .setBackground(COULEUR_ANNEE[annee] || "#182433")
    .setFontColor("#ffffff").setFontSize(19).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  let rCurseur = rowStart + 1;
  for (let r = 0; r < GRID_ROWS; r++) {
    const m1 = r * 2;
    const m2 = r * 2 + 1;
    const n1 = compteurs[`${annee}_${m1}`] || 0;
    const n2 = compteurs[`${annee}_${m2}`] || 0;
    const nbLignesPair = Math.max(MIN_LIGNES_PAR_MOIS, Math.max(n1, n2) + MARGE_LIGNES_VIDES);

    _creerBlocMoisDynamique(sheet, m1, rCurseur, nbLignesPair);
    _creerBlocMoisDynamique(sheet, m2, rCurseur, nbLignesPair);

    if (blocInfoMap) {
      blocInfoMap[`${annee}_${m1}`] = { ligneData: rCurseur + 2, colDebut: 1, nbLignesData: nbLignesPair };
      blocInfoMap[`${annee}_${m2}`] = { ligneData: rCurseur + 2, colDebut: 1 + BLOC_WIDTH, nbLignesData: nbLignesPair };
    }
    rCurseur += 2 + nbLignesPair + GRID_GAP_ROWS;
  }

  rCurseur += GAP_AVANT_INDET;
  const nbIndetNecessaire = Math.max(LIGNES_INDETERMINE, (compteurs[`${annee}_indet`] || 0) + MARGE_LIGNES_VIDES);
  _creerBlocIndetermineDynamique(sheet, annee, rCurseur, nbIndetNecessaire);
  if (blocInfoMap) blocInfoMap[`${annee}_indet`] = { ligneData: rCurseur + 2, colDebut: 1, nbLignesData: nbIndetNecessaire };

  rCurseur += 2 + nbIndetNecessaire;
  sheet.setRowHeight(rCurseur, 10);
  sheet.getRange(rCurseur, 1, 1, LARGEUR_GRILLE).setBackground("#E63946");

  return rCurseur + 1;
}

function _injecterSauvegardeEnBloc(sheet, sauvegarde, blocInfoMap) {
  const groupes = {};
  const titresVus = {};
  const echecs = [];

  sauvegarde.forEach(d => {
    const nom = (d.titre || "").toString().trim();
    if (!nom) return;

    const parsed = _parseDateEntree(d.date);
    let key, ligne;

    if (parsed.type === "annee" || parsed.type === "saison") {
      key = `${parsed.annee}_indet`;
      ligne = {
        periode: parsed.type === "saison" ? `${parsed.saison} ${parsed.annee}` : `${parsed.annee}`,
        titre: nom, cat: d.cat || CATEGORIES[0], plateforme: d.plateforme || "",
        lien: d.lien || "", statut: (d.statut && d.statut.toString().trim()) ? d.statut.toString().trim() : "❌",
        sortKey: parsed.type === "saison" ? SAISON_ORDRE[parsed.saison] : 99
      };
    } else if (parsed.type === "mois") {
      key = `${parsed.annee}_${parsed.mois - 1}`;
      const jourNum = (parsed.jour === "?" || parsed.jour === "") ? 99 : parseInt(parsed.jour, 10);
      ligne = {
        jour: jourNum === 99 ? "?" : String(jourNum).padStart(2, "0"),
        titre: nom, cat: d.cat || CATEGORIES[0], plateforme: d.plateforme || "",
        lien: d.lien || "", statut: (d.statut && d.statut.toString().trim()) ? d.statut.toString().trim() : "❌", sortKey: jourNum
      };
    } else {
      echecs.push(`${nom} (Date invalide: ${d.date})`);
      return;
    }

    if (!titresVus[key]) titresVus[key] = new Set();
    if (titresVus[key].has(nom.toLowerCase().trim())) {
      echecs.push(`${nom}`);
      return;
    }
    titresVus[key].add(nom.toLowerCase().trim());

    if (!groupes[key]) groupes[key] = [];
    groupes[key].push(ligne);
  });

  Object.keys(groupes).forEach(key => {
    const info = blocInfoMap[key];
    if (!info) return;

    let lignes = groupes[key].sort((a, b) => a.sortKey - b.sortKey);
    const estIndet = key.endsWith("_indet");

    const vals = lignes.map(l => [estIndet ? l.periode : l.jour, _neutraliserFormule(l.titre), l.cat, _neutraliserFormule(l.plateforme), _neutraliserFormule(l.lien), l.statut]);
    while (vals.length < info.nbLignesData) vals.push(["", "", "", "", "", ""]);

    sheet.getRange(info.ligneData, info.colDebut, info.nbLignesData, COLS_PAR_MOIS).setValues(vals.slice(0, info.nbLignesData));

    if (estIndet) {
      _stylerBlocIndetermine(sheet, { ligneData: info.ligneData, nbLignesData: info.nbLignesData });
    } else {
      _stylerBlocMois(sheet, { ligneData: info.ligneData, colDebut: info.colDebut, nbLignesData: info.nbLignesData }, parseInt(key.split("_")[1], 10));
    }
  });

  return { total: sauvegarde.length, echecs: echecs.length, detailsEchecs: echecs };
}

function _creerBlocMoisDynamique(sheet, m, rowStart, nbLignes) {
  const colDebut = 1 + (m % GRID_COLS) * BLOC_WIDTH;
  const couleur = _couleurMois(m);

  sheet.setRowHeight(rowStart, 36);
  sheet.getRange(rowStart, colDebut, 1, COLS_PAR_MOIS).merge().setValue(MOIS_NOMS[m].toUpperCase()).setBackground(couleur).setFontColor("#ffffff").setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");

  sheet.setRowHeight(rowStart + 1, 24);
  sheet.getRange(rowStart + 1, colDebut, 1, COLS_PAR_MOIS).setValues([["JOUR","TITRE","CATÉGORIE","OÙ VOIR","LIEN","STATUT"]]).setBackground(_assombrir(couleur)).setFontColor("#ffffff").setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center").setVerticalAlignment("middle");

  const lineDataStart = rowStart + 2;
  sheet.setRowHeights(lineDataStart, nbLignes, 32);

  const bgRows = [];
  for (let r = 0; r < nbLignes; r++) bgRows.push(_ligneVide(r % 2 === 0 ? "#FFFFFF" : "#F6F7F8"));
  sheet.getRange(lineDataStart, colDebut, nbLignes, COLS_PAR_MOIS).setBackgrounds(bgRows);

  sheet.getRange(lineDataStart, colDebut + OFF_JOUR, nbLignes, 1).setHorizontalAlignment("center").setFontSize(11);
  sheet.getRange(lineDataStart, colDebut + OFF_TITRE, nbLignes, 1).setFontSize(11);
  sheet.getRange(lineDataStart, colDebut + OFF_CAT, nbLignes, 1).setHorizontalAlignment("center").setFontSize(10);
  sheet.getRange(lineDataStart, colDebut + OFF_PLATEFORME, nbLignes, 1).setHorizontalAlignment("center").setFontSize(10);
  sheet.getRange(lineDataStart, colDebut + OFF_LIEN, nbLignes, 1).setHorizontalAlignment("left").setFontSize(9);
  sheet.getRange(lineDataStart, colDebut + OFF_STATUT, nbLignes, 1).setHorizontalAlignment("center").setFontSize(13);

  _appliquerValidationCategorie(sheet.getRange(lineDataStart, colDebut + OFF_CAT, nbLignes, 1));
  _appliquerValidationPlateforme(sheet.getRange(lineDataStart, colDebut + OFF_PLATEFORME, nbLignes, 1));

  sheet.getRange(rowStart, colDebut, 2 + nbLignes, COLS_PAR_MOIS).setBorder(true,true,true,true,false,false,"#D9DEE3",SpreadsheetApp.BorderStyle.SOLID);
}

function _creerBlocIndetermineDynamique(sheet, annee, rowStart, nbLignes) {
  const nbL = nbLignes || LIGNES_INDETERMINE;
  sheet.setRowHeight(rowStart, 30);
  sheet.getRange(rowStart, 1, 1, LARGEUR_GRILLE).merge().setValue(`📆 DATE INDÉTERMINÉE / SAISON   (${annee})`).setBackground(COULEUR_INDETERMINE).setFontColor("#ffffff").setFontSize(12).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");

  sheet.setRowHeight(rowStart + 1, 22);
  sheet.getRange(rowStart + 1, 1, 1, COLS_PAR_MOIS).setValues([["PÉRIODE","TITRE","CATÉGORIE","OÙ VOIR","LIEN","STATUT"]]).setBackground(_assombrir(COULEUR_INDETERMINE)).setFontColor("#ffffff").setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center").setVerticalAlignment("middle");

  const lineDataStart = rowStart + 2;
  sheet.setRowHeights(lineDataStart, nbL, 26);

  const bgRowsI = [];
  for (let r = 0; r < nbL; r++) bgRowsI.push(_ligneVide(r % 2 === 0 ? "#FFFFFF" : "#F6F7F8"));
  sheet.getRange(lineDataStart, 1, nbL, COLS_PAR_MOIS).setBackgrounds(bgRowsI);

  sheet.getRange(lineDataStart, 1 + OFF_JOUR, nbL, 1).setHorizontalAlignment("center").setFontSize(10);
  sheet.getRange(lineDataStart, 1 + OFF_TITRE, nbL, 1).setFontSize(10);
  sheet.getRange(lineDataStart, 1 + OFF_CAT, nbL, 1).setHorizontalAlignment("center").setFontSize(10);
  sheet.getRange(lineDataStart, 1 + OFF_PLATEFORME, nbL, 1).setHorizontalAlignment("center").setFontSize(10);
  sheet.getRange(lineDataStart, 1 + OFF_LIEN, nbL, 1).setHorizontalAlignment("left").setFontSize(9);
  sheet.getRange(lineDataStart, 1 + OFF_STATUT, nbL, 1).setHorizontalAlignment("center").setFontSize(13);

  _appliquerValidationCategorie(sheet.getRange(lineDataStart, 1 + OFF_CAT, nbL, 1));
  _appliquerValidationPlateforme(sheet.getRange(lineDataStart, 1 + OFF_PLATEFORME, nbL, 1));

  sheet.getRange(rowStart, 1, 2 + nbL, COLS_PAR_MOIS).setBorder(true,true,true,true,false,false,"#D9DEE3",SpreadsheetApp.BorderStyle.SOLID);
}

function _sauvegarderDonneesGenerique(sheet) {
  const maxRow = sheet.getMaxRows();
  const maxCol = sheet.getMaxColumns();
  if (maxRow === 0 || maxCol === 0) return [];

  const values = sheet.getRange(1, 1, maxRow, maxCol).getValues();
  const anchors = [];

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (typeof values[r][c] !== "string") continue;
      const txtNorm = _normSimple(values[r][c]);
      if (txtNorm === "") continue;

      const mAnnee = txtNorm.match(/sorties\s+(\d{4})/);
      if (mAnnee) { anchors.push({ row: r, type: "annee", annee: parseInt(mAnnee[1], 10), col: c }); continue; }

      const idxMois = MOIS_NOMS.findIndex(m => _normSimple(m) === txtNorm);
      if (idxMois !== -1) { anchors.push({ row: r, type: "mois", mois: idxMois, col: c }); continue; }

      if (txtNorm.includes("date indeterminee")) anchors.push({ row: r, type: "indet", col: c });
    }
  }
  anchors.sort((a, b) => a.row - b.row);

  let anneeCourante = null;
  const blocs = [];

  anchors.forEach(a => {
    if (a.type === "annee") { anneeCourante = a.annee; return; }
    blocs.push({ row: a.row, col: a.col, type: a.type, mois: a.mois, annee: anneeCourante });
  });

  const lignesAncrages = anchors.map(a => a.row).sort((x, y) => x - y);
  function _finDeBloc(ligneTitre) {
    for (let i = 0; i < lignesAncrages.length; i++) {
      if (lignesAncrages[i] > ligneTitre) return lignesAncrages[i] - 1;
    }
    return values.length - 1;
  }

  const sauvegarde = [];
  blocs.forEach(b => {
    if (b.annee === null || !YEARS.includes(b.annee)) return;

    const ligneDataDebut = b.row + 2;
    const ligneDataFin = _finDeBloc(b.row);
    if (ligneDataFin < ligneDataDebut) return;

    for (let r = ligneDataDebut; r <= ligneDataFin; r++) {
      if (!values[r]) break;
      const jourOuPeriode = (values[r][b.col + 0] || "").toString().trim();
      const titre = (values[r][b.col + 1] || "").toString().trim();
      const cat = (values[r][b.col + 2] || "").toString().trim();

      if (titre === "") continue;

      const plateforme = (values[r][b.col + 3] || "").toString().trim();
      const lien = (values[r][b.col + 4] || "").toString().trim();
      const statut = (values[r][b.col + 5] || "").toString().trim();

      if (b.type === "indet") {
        sauvegarde.push({ titre, date: jourOuPeriode || `${b.annee}`, cat: cat || CATEGORIES[0], plateforme, lien, statut: statut || "❌" });
      } else {
        sauvegarde.push({
          titre,
          date: `${(jourOuPeriode==="?"||jourOuPeriode==="")?"?":jourOuPeriode.padStart(2,"0")}/${String(b.mois+1).padStart(2,"0")}/${b.annee}`,
          cat: cat || CATEGORIES[0], plateforme, lien, statut: statut || "❌"
        });
      }
    }
  });
  return sauvegarde;
}

function ajouterSortieDepuisHTML(data) {
  _exigerContexteTableur();
  const msg = insererEntree_(data.nom, data.date, data.categorie, data.plateforme, data.lien, data.vu);
  if (msg.startsWith("✅")) {
    enregistrerAjoutRecent_({ titre: data.nom.toString().trim(), dateStr: data.date.toString().trim() });
  }
  mettreEnValeurProchaineSortie();
  return msg;
}

function _stylerBlocMois(sheet, info, numMois, valsExistants) {
  if (info.ligneData === -1) return;
  const range = sheet.getRange(info.ligneData, info.colDebut, info.nbLignesData, COLS_PAR_MOIS);
  const vals = valsExistants || range.getValues();
  const couleurMois = _couleurMois(numMois);
  const bgs = [], fcs = [], fws = [];

  vals.forEach((row, r) => {
    const titre = (row[OFF_TITRE] || "").toString().trim();
    const bgRow = r % 2 === 0 ? "#ffffff" : "#f6f7f8";
    if (titre) {
      const bgCat = _bgCouleurCat(row[OFF_CAT]);
      bgs.push([couleurMois, bgCat, bgCat, bgCat, bgCat, bgCat]);
      fcs.push(["#ffffff", "#1a1a2e", "#20252b", "#20252b", "#20252b", "#20252b"]);
      fws.push(["bold", "bold", "normal", "normal", "normal", "normal"]);
    } else {
      bgs.push(Array(COLS_PAR_MOIS).fill(bgRow));
      fcs.push(Array(COLS_PAR_MOIS).fill("#20252b"));
      fws.push(Array(COLS_PAR_MOIS).fill("normal"));
    }
  });
  range.setBackgrounds(bgs);
  range.setFontColors(fcs);
  range.setFontWeights(fws);
}

function _stylerBlocIndetermine(sheet, info, valsExistants) {
  if (info.ligneData === -1) return;
  const range = sheet.getRange(info.ligneData, 1, info.nbLignesData, COLS_PAR_MOIS);
  const vals = valsExistants || range.getValues();
  const bgs = [], fcs = [], fws = [];

  vals.forEach((row, r) => {
    const titre = (row[OFF_TITRE] || "").toString().trim();
    const bgRow = r % 2 === 0 ? "#ffffff" : "#f6f7f8";
    if (titre) {
      const bgCat = _bgCouleurCat(row[OFF_CAT]);
      bgs.push([COULEUR_INDETERMINE, bgCat, bgCat, bgCat, bgCat, bgCat]);
      fcs.push(["#ffffff", "#1a1a2e", "#20252b", "#20252b", "#20252b", "#20252b"]);
      fws.push(["bold", "bold", "normal", "normal", "normal", "normal"]);
    } else {
      bgs.push(Array(COLS_PAR_MOIS).fill(bgRow));
      fcs.push(Array(COLS_PAR_MOIS).fill("#20252b"));
      fws.push(Array(COLS_PAR_MOIS).fill("normal"));
    }
  });
  range.setBackgrounds(bgs);
  range.setFontColors(fcs);
  range.setFontWeights(fws);
}

// ============================================================
//   INSERTION, SUPPRESSION & REMPLISSAGE STATUT ❌
// ============================================================

function insererEntree_(nom, dateStr, cat, plateforme, lien, statut) {
  nom = (nom || "").toString().trim();
  if (!nom) return "❌ Le nom est vide.";

  cat = (cat && cat.toString().trim()) ? cat.toString().trim() : CATEGORIES[0];
  plateforme = (plateforme || "").toString().trim();
  lien = (lien || "").toString().trim();
  statut = (statut && statut.toString().trim()) ? statut.toString().trim() : "❌";

  const parsed = _parseDateEntree(dateStr);
  if (parsed.type === "invalide") return `❌ Date invalide : "${dateStr}" (${parsed.raison || "Format non reconnu"})`;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOM_FEUILLE);
  if (!sheet) return `❌ Onglet "${NOM_FEUILLE}" introuvable.`;

  if (parsed.type === "annee" || parsed.type === "saison") return _insererEntreeIndetermine(sheet, nom, parsed, cat, plateforme, lien, statut);

  const jourNum = (parsed.jour === "?" || parsed.jour === "") ? 99 : parseInt(parsed.jour, 10);
  let info = _infoMois(sheet, parsed.annee, parsed.mois - 1);
  if (info.ligneData === -1) return `❌ Bloc introuvable pour ${MOIS_NOMS[parsed.mois - 1]} ${parsed.annee}.`;

  const colDebut = info.colDebut;
  let range = sheet.getRange(info.ligneData, colDebut, info.nbLignesData, COLS_PAR_MOIS);
  let vals = range.getValues();

  for (let i = 0; i < vals.length; i++) {
    if (vals[i][OFF_TITRE].toString().toLowerCase().trim() === nom.toLowerCase()) {
      return `⚠️ DOUBLON : "${nom}" existe déjà.`;
    }
  }

  let ligneInsert = -1, dernierRempli = -1;
  for (let i = 0; i < vals.length; i++) {
    const t = vals[i][OFF_TITRE];
    if (t !== "" && t !== null && t !== undefined) {
      dernierRempli = i;
      const v = vals[i][OFF_JOUR];
      const jourEx = (v === "?" || v === "") ? 99 : parseInt(v, 10);
      if (ligneInsert === -1 && jourNum <= jourEx) ligneInsert = i;
    }
  }

  if (dernierRempli >= info.nbLignesData - 1) {
    const ligneDernierData = info.ligneData + info.nbLignesData - 1;
    sheet.insertRowAfter(ligneDernierData);
    sheet.setRowHeight(ligneDernierData + 1, 32);

    for (let c = 0; c < GRID_COLS; c++) {
      const colD = 1 + c * BLOC_WIDTH;
      _appliquerValidationCategorie(sheet.getRange(ligneDernierData + 1, colD + OFF_CAT, 1, 1));
      _appliquerValidationPlateforme(sheet.getRange(ligneDernierData + 1, colD + OFF_PLATEFORME, 1, 1));
      sheet.getRange(info.ligneTitreMois, colD, 2 + info.nbLignesData + 1, COLS_PAR_MOIS).setBorder(true,true,true,true,false,false,"#D9DEE3",SpreadsheetApp.BorderStyle.SOLID);
    }
    info.nbLignesData += 1;
    range = sheet.getRange(info.ligneData, colDebut, info.nbLignesData, COLS_PAR_MOIS);
    vals = range.getValues();
  }

  if (ligneInsert === -1) ligneInsert = dernierRempli + 1;
  if (ligneInsert < 0) ligneInsert = 0;

  vals.splice(ligneInsert, 0, [jourNum === 99 ? "?" : String(jourNum).padStart(2,"0"), _neutraliserFormule(nom), cat, _neutraliserFormule(plateforme), _neutraliserFormule(lien), statut]);
  if (vals.length > info.nbLignesData) vals.pop();

  range.setValues(vals);
  _stylerBlocMois(sheet, info, parsed.mois - 1, vals);
  _invaliderCacheJSON();
  return `✅ "${nom}" ajouté !`;
}

function _insererEntreeIndetermine(sheet, nom, parsed, cat, plateforme, lien, statut) {
  const annee = parsed.annee;
  let info = _infoIndetermine(sheet, annee);
  if (info.ligneData === -1) return `❌ Bloc indéterminé introuvable pour ${annee}.`;

  const periodeAff = parsed.type === "saison" ? `${parsed.saison} ${annee}` : `${annee}`;
  const cleSort = parsed.type === "saison" ? SAISON_ORDRE[parsed.saison] : 99;

  let range = sheet.getRange(info.ligneData, 1, info.nbLignesData, COLS_PAR_MOIS);
  let vals = range.getValues();

  for (let i = 0; i < vals.length; i++) {
    if (vals[i][OFF_TITRE].toString().toLowerCase().trim() === nom.toLowerCase().trim()) return `⚠️ DOUBLON en indéterminé.`;
  }

  let ligneInsert = -1, dernierRempli = -1;
  for (let i = 0; i < vals.length; i++) {
    const t = vals[i][OFF_TITRE];
    if (t !== "" && t !== null && t !== undefined) {
      dernierRempli = i;
      const cleEx = _cleSortPeriode(vals[i][OFF_JOUR]);
      if (ligneInsert === -1 && cleSort <= cleEx) ligneInsert = i;
    }
  }

  if (dernierRempli >= info.nbLignesData - 1) {
    const ligneDernierData = info.ligneData + info.nbLignesData - 1;
    sheet.insertRowAfter(ligneDernierData);
    sheet.setRowHeight(ligneDernierData + 1, 26);

    _appliquerValidationCategorie(sheet.getRange(ligneDernierData + 1, 1 + OFF_CAT, 1, 1));
    _appliquerValidationPlateforme(sheet.getRange(ligneDernierData + 1, 1 + OFF_PLATEFORME, 1, 1));
    sheet.getRange(info.ligneTitre, 1, 2 + info.nbLignesData + 1, COLS_PAR_MOIS).setBorder(true,true,true,true,false,false,"#D9DEE3",SpreadsheetApp.BorderStyle.SOLID);

    info.nbLignesData += 1;
    range = sheet.getRange(info.ligneData, 1, info.nbLignesData, COLS_PAR_MOIS);
    vals = range.getValues();
  }

  if (ligneInsert === -1) ligneInsert = dernierRempli + 1;
  if (ligneInsert < 0) ligneInsert = 0;

  vals.splice(ligneInsert, 0, [periodeAff, _neutraliserFormule(nom), cat, _neutraliserFormule(plateforme), _neutraliserFormule(lien), statut]);
  if (vals.length > info.nbLignesData) vals.pop();

  range.setValues(vals);
  _stylerBlocIndetermine(sheet, info, vals);
  _invaliderCacheJSON();
  return `✅ "${nom}" ajouté (date indéterminée)`;
}

// Remplit avec ❌ le statut de toutes les sorties dont le statut est vide
function remplirCasesVidesParCroix() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOM_FEUILLE);
  if (!sheet) return _alerte(`❌ Onglet "${NOM_FEUILLE}" introuvable.`);

  const fullData = sheet.getDataRange().getValues();
  const vides = _parcourirEntrees(fullData, e => !e.statut);

  vides.forEach(e => {
    sheet.getRange(e.ligneAbs, e.colDebut + OFF_STATUT).setValue("❌");
  });

  _invaliderCacheJSON();
  _alerte(`✅ ${vides.length} statut(s) vide(s) remplacé(s) par ❌.`);
}

function supprimerEntreeInterne_(ss, cible, archiver) {
  if (archiver) {
    const horodatage = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    const dateRestaurable = cible.indetermine ?
      (cible.periodeActuelle || `${cible.annee}`) :
      `${(cible.jourActuel && cible.jourActuel !== "?") ? cible.jourActuel : "?"}/${cible.mois + 1}/${cible.annee}`;
    const fC = _getOrCreateCorbeille(ss);
    const rowC = fC.getLastRow() + 1;
    fC.getRange(rowC, 2).setNumberFormat("@");
    fC.getRange(rowC, 1, 1, 7).setValues([[cible.titre, dateRestaurable, cible.cat, cible.plateforme, cible.lien, cible.statut, horodatage]]);
  }

  const sheet = ss.getSheetByName(NOM_FEUILLE);
  const infoRef = cible.indetermine ? _infoIndetermine(sheet, cible.annee) : _infoMois(sheet, cible.annee, cible.mois);
  const offsetLigne = cible.ligneAbs - infoRef.ligneData;

  const range = sheet.getRange(infoRef.ligneData, cible.colDebut, infoRef.nbLignesData, COLS_PAR_MOIS);
  const vals = range.getValues();

  if (offsetLigne >= 0 && offsetLigne < vals.length) {
    vals.splice(offsetLigne, 1);
    vals.push(["", "", "", "", "", ""]);
    range.setValues(vals);
  }

  if (cible.indetermine) _stylerBlocIndetermine(sheet, infoRef, vals);
  else _stylerBlocMois(sheet, infoRef, cible.mois, vals);

  _invaliderCacheJSON();
}

// ============================================================
//   GESTION PROCHAINE SORTIE & SORTIES DU JOUR (J-0)
// ============================================================

function mettreEnValeurProchaineSortie() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOM_FEUILLE);
  if (!sheet) return;
  _invaliderCacheJSON();

  const props = PropertiesService.getScriptProperties();
  const anciensHighlightsStr = props.getProperty(PROP_KEY_HIGHLIGHTS);
  if (anciensHighlightsStr) {
    try {
      const anciens = JSON.parse(anciensHighlightsStr);
      anciens.forEach(a => {
        if (a.ligne && a.col) {
          sheet.getRange(a.ligne, a.col, 1, COLS_PAR_MOIS).setBorder(false, false, false, false, false, false);
        }
        if (a.indet) {
          _stylerBlocIndetermine(sheet, _infoIndetermine(sheet, a.annee));
        } else {
          _stylerBlocMois(sheet, _infoMois(sheet, a.annee, a.mois), a.mois);
        }
      });
    } catch(e){}
  }

  const aujourdhui = new Date();
  aujourdhui.setHours(0, 0, 0, 0);
  const candidats = [];
  const fullData = sheet.getDataRange().getValues();

  YEARS.forEach(annee => {
    for (let m = 0; m < 12; m++) {
      const info = _infoMois(fullData, annee, m);
      if (info.ligneData === -1) continue;
      const rStart = info.ligneData - 1;
      const cStart = info.colDebut - 1;

      for (let r = 0; r < info.nbLignesData; r++) {
        if (!fullData[rStart + r]) continue;
        const titre = (fullData[rStart + r][cStart + OFF_TITRE] || "").toString().trim();
        const jourV = (fullData[rStart + r][cStart + OFF_JOUR] || "").toString().trim();

        if (titre && jourV !== "?" && jourV !== "") {
          const jour = parseInt(jourV, 10);
          if (!isNaN(jour)) {
            const dateSortie = new Date(annee, m, jour);
            dateSortie.setHours(0, 0, 0, 0);
            if (dateSortie >= aujourdhui) {
              candidats.push({ date: dateSortie, ligne: info.ligneData + r, col: info.colDebut, titre, annee, mois: m, indet: false });
            }
          }
        }
      }
    }
  });

  if (candidats.length === 0) {
    sheet.getRange(4, 1).setValue("⭐ Prochaine sortie : —");
    sheet.getRange(4, 1, 2, LARGEUR_GRILLE).setBackground("#FFF3CD").setFontColor("#765600");
    props.deleteProperty(PROP_KEY_HIGHLIGHTS);
    return;
  }

  candidats.sort((a, b) => a.date - b.date);
  const dateCible = candidats[0].date.getTime();
  const prochaines = candidats.filter(c => c.date.getTime() === dateCible);
  const estAujourdhui = (dateCible === aujourdhui.getTime());

  const jourFormate = String(candidats[0].date.getDate()).padStart(2, "0");
  const moisFormate = String(candidats[0].date.getMonth() + 1).padStart(2, "0");
  const anneeFormate = candidats[0].date.getFullYear();

  if (estAujourdhui) {
    prochaines.forEach(c => {
      sheet.getRange(c.ligne, c.col, 1, COLS_PAR_MOIS).setBackground("#E8F5E9").setBorder(true, true, true, true, false, false, "#2E7D32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      sheet.getRange(c.ligne, c.col + OFF_JOUR).setBackground("#2E7D32").setFontColor("#FFFFFF").setFontWeight("bold");
      sheet.getRange(c.ligne, c.col + OFF_TITRE).setFontWeight("bold").setFontColor("#1B5E20");
    });

    const headerText = `🎉 SORTIE(S) DU JOUR (${jourFormate}/${moisFormate}) : ${prochaines.map(c => c.titre).join(", ")}`;
    sheet.getRange(4, 1).setValue(headerText);
    sheet.getRange(4, 1, 2, LARGEUR_GRILLE).setBackground("#D1E7DD").setFontColor("#0F5132");
  } else {
    prochaines.forEach(c => {
      sheet.getRange(c.ligne, c.col, 1, COLS_PAR_MOIS).setBackground("#FFF8C5").setBorder(true, true, true, true, false, false, "#F0B429", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      sheet.getRange(c.ligne, c.col + OFF_JOUR).setBackground("#F0B429").setFontColor("#1A1A2E").setFontWeight("bold");
      sheet.getRange(c.ligne, c.col + OFF_TITRE).setFontWeight("bold").setFontColor("#7D4E00");
    });

    const diffJours = Math.round((candidats[0].date - aujourdhui) / 86400000);
    const headerText = `⭐ Prochaine sortie (dans ${diffJours} j. • ${jourFormate}/${moisFormate}/${anneeFormate}) : ${prochaines.map(c => c.titre).join(", ")}`;
    sheet.getRange(4, 1).setValue(headerText);
    sheet.getRange(4, 1, 2, LARGEUR_GRILLE).setBackground("#FFF3CD").setFontColor("#765600");
  }

  props.setProperty(PROP_KEY_HIGHLIGHTS, JSON.stringify(prochaines.map(p => ({ ligne: p.ligne, col: p.col, annee: p.annee, mois: p.mois, indet: p.indet }))));
}

function _choisirParmiResultats(ui, resultats) {
  if (resultats.length === 0) return null;
  if (resultats.length === 1) return resultats[0];

  const liste = resultats.map((r, i) => `${i + 1}. ${r.titre}  —  ${r.indetermine ? r.periodeActuelle + ' ' + r.annee : MOIS_NOMS[r.mois] + ' ' + r.annee + ' (jour: ' + r.jourActuel + ')'}`).join("\n");
  const repChoix = ui.prompt("Plusieurs résultats trouvés", `Indique le numéro souhaité :\n\n${liste}`, ui.ButtonSet.OK_CANCEL);

  if (repChoix.getSelectedButton() !== ui.Button.OK) return null;
  const idx = parseInt(repChoix.getResponseText().trim(), 10) - 1;

  if (isNaN(idx) || idx < 0 || idx >= resultats.length) { ui.alert("❌ Numéro invalide."); return null; }
  return resultats[idx];
}

function rechercherSortieRapide() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOM_FEUILLE);
  const rep = ui.prompt("🔎 Recherche globale", "Nom :", ui.ButtonSet.OK_CANCEL);
  if (rep.getSelectedButton() !== ui.Button.OK) return;

  const recherche = rep.getResponseText().trim().toLowerCase();
  const res = _parcourirEntrees(sheet, e => e.titre.toLowerCase().includes(recherche));

  if (res.length === 0) return ui.alert(`❌ Aucun résultat pour "${recherche}".`);

  const liste = res.map(r => `• ${r.titre}  —  ${r.indetermine ? r.periodeActuelle + ' ' + r.annee : r.jourActuel + '/' + (r.mois+1) + '/' + r.annee}  —  ${r.cat}  ${r.statut}`).join("\n");
  ui.alert(`🔎 ${res.length} résultat(s)`, liste, ui.ButtonSet.OK);
}

function modifierDate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOM_FEUILLE);
  const ui = SpreadsheetApp.getUi();

  const repNom = ui.prompt("✏️ Modifier une date", "Nom de l'œuvre :", ui.ButtonSet.OK_CANCEL);
  if (repNom.getSelectedButton() !== ui.Button.OK) return;

  const cible = _choisirParmiResultats(ui, _parcourirEntrees(sheet, e => e.titre.toLowerCase().includes(repNom.getResponseText().trim().toLowerCase())));
  if (!cible) return;

  const repDate = ui.prompt(`✏️ Nouvelle date pour "${cible.titre}"`, "JJ/MM/AAAA, MM/AAAA ou Saison AAAA...", ui.ButtonSet.OK_CANCEL);
  if (repDate.getSelectedButton() !== ui.Button.OK) return;

  const nvDate = repDate.getResponseText().trim();
  if (!nvDate) return ui.alert("❌ Date vide.");

  const ancDate = cible.indetermine ? (cible.periodeActuelle || `${cible.annee}`) : `${cible.jourActuel || "?"}/${cible.mois + 1}/${cible.annee}`;
  supprimerEntreeInterne_(ss, cible, false);

  const msg = insererEntree_(cible.titre, nvDate, cible.cat, cible.plateforme, cible.lien, cible.statut);
  if (!msg.startsWith("✅")) {
    insererEntree_(cible.titre, ancDate, cible.cat, cible.plateforme, cible.lien, cible.statut);
    return ui.alert(`${msg}\n\n↩️ Remis à l'origine (${ancDate}).`);
  }

  mettreEnValeurProchaineSortie();
  ui.alert(msg);
}

function marquerStatutSortie() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOM_FEUILLE);
  const ui = SpreadsheetApp.getUi();

  const repNom = ui.prompt("✅ Statut", "Nom :", ui.ButtonSet.OK_CANCEL);
  if (repNom.getSelectedButton() !== ui.Button.OK) return;

  const cible = _choisirParmiResultats(ui, _parcourirEntrees(sheet, e => e.titre.toLowerCase().includes(repNom.getResponseText().trim().toLowerCase())));
  if (!cible) return;

  const rep = ui.prompt(`Statut de "${cible.titre}"`, "1 = ✔️ Vu\n2 = ❌ Pas vu\n3 = Effacer", ui.ButtonSet.OK_CANCEL);
  if (rep.getSelectedButton() !== ui.Button.OK) return;

  const val = {"1":"✔️", "2":"❌"}[rep.getResponseText().trim()] || "";
  sheet.getRange(cible.ligneAbs, cible.colDebut + OFF_STATUT).setValue(val);

  _invaliderCacheJSON();
  ui.alert("✅ Statut mis à jour !");
}

function supprimerSortie() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOM_FEUILLE);
  const ui = SpreadsheetApp.getUi();

  const repNom = ui.prompt("🗑️ Supprimer", "Nom :", ui.ButtonSet.OK_CANCEL);
  if (repNom.getSelectedButton() !== ui.Button.OK) return;

  const cible = _choisirParmiResultats(ui, _parcourirEntrees(sheet, e => e.titre.toLowerCase().includes(repNom.getResponseText().trim().toLowerCase())));
  if (!cible) return;

  supprimerEntreeInterne_(ss, cible, true);
  mettreEnValeurProchaineSortie();
  ui.alert(`✅ "${cible.titre}" déplacé dans la corbeille.`);
}

function restaurerSortie() {
  const ui = SpreadsheetApp.getUi();
  const fCorbeille = _getOrCreateCorbeille(SpreadsheetApp.getActiveSpreadsheet());
  const data = fCorbeille.getDataRange().getValues();

  if (data.length <= 1) return ui.alert("♻️ Corbeille vide.");

  const rep = ui.prompt("♻️ Restaurer", `Indique le numéro :\n\n${data.slice(1).map((r, i) => `${i + 1}. ${r[0]} (${r[1]})`).join("\n")}`, ui.ButtonSet.OK_CANCEL);
  if (rep.getSelectedButton() !== ui.Button.OK) return;

  const idx = parseInt(rep.getResponseText().trim(), 10) - 1;
  if (isNaN(idx) || idx < 0 || idx >= data.length - 1) return ui.alert("❌ Numéro invalide.");

  const row = data[idx + 1];
  const msg = insererEntree_(row[0], row[1], row[2], row[3], row[4], row[5]);

  if (!msg.startsWith("✅")) return ui.alert(`❌ Erreur lors de la restauration : ${msg}`);

  fCorbeille.deleteRow(idx + 2);
  mettreEnValeurProchaineSortie();
  ui.alert(msg);
}

// ============================================================
//   MOTEUR D'ENRICHISSEMENT OPTIMISÉ (Batch Fetching)
// ============================================================

function enrichirAuChoix(choixPre) {
  const ui = SpreadsheetApp.getUi();
  let choix = (typeof choixPre === "string") ? choixPre : null;

  if (!choix) {
    const repChoix = ui.prompt("🚀 Enrichissement global", "1 = 🍿 Animés\n2 = 🖼️ Films & Séries\n3 = 🎮 Jeux\n4 = 🔥 TOUT", ui.ButtonSet.OK_CANCEL);
    if (repChoix.getSelectedButton() !== ui.Button.OK) return;
    choix = repChoix.getResponseText().trim();
  }

  const repForcer = ui.alert("⚙️ Mode", "Réécrire 'Où voir' et 'Lien' même s'ils sont déjà remplis ?", ui.ButtonSet.YES_NO_CANCEL);
  if (repForcer === ui.Button.CANCEL) return;
  const forcer = repForcer === ui.Button.YES;

  const props = PropertiesService.getScriptProperties();
  const tmdbKey = props.getProperty(PROP_KEY_TMDB);
  const igdbId = props.getProperty(PROP_KEY_IGDB_CLIENT_ID);
  const igdbSecret = props.getProperty(PROP_KEY_IGDB_CLIENT_SECRET);

  let rapport = [];

  if (["1", "4"].includes(choix)) {
    const r = _enrichirAnimesCore(forcer);
    rapport.push(`🍿 Animés: ${r.majPlat} plateformes, ${r.majLien} liens`);
  }

  if (["2", "4"].includes(choix)) {
    if (!tmdbKey) ui.alert("🔑 Clé TMDB manquante.");
    else {
      const r = _enrichirTMDBCore(tmdbKey, forcer);
      rapport.push(`🖼️ Films/Séries: ${r.majPlat} plateformes, ${r.majLien} liens`);
    }
  }

  if (["3", "4"].includes(choix)) {
    if (!igdbId || !igdbSecret) ui.alert("🔑 Clés IGDB manquantes.");
    else {
      const tok = _obtenirTokenIGDB(igdbId, igdbSecret);
      if (tok) {
        const r = _enrichirIGDBCore(igdbId, tok, forcer);
        rapport.push(`🎮 Jeux: ${r.majPlat} plateformes, ${r.majLien} liens`);
      }
    }
  }

  _invaliderCacheJSON();
  SpreadsheetApp.flush();
  ui.alert("✅ Terminé !\n" + rapport.join("\n"));
}

function _enrichirAnimesCore(forcer) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOM_FEUILLE);
  const fullData = sheet.getDataRange().getValues();
  const entries = _parcourirEntrees(fullData, e => _estAnime(e.cat) && (forcer || !e.plateforme || !e.lien));
  if (entries.length === 0) return { majPlat: 0, majLien: 0, echecs: 0 };

  let majPlat = 0, majLien = 0, echecs = 0;

  const requests = entries.map(e => {
    const q = _extraireNomBaseAnime(e.titre);
    return {
      url: "https://graphql.anilist.co",
      method: "post",
      contentType: "application/json",
      muteHttpExceptions: true,
      payload: JSON.stringify({
        query: `query ($search: String) {
          Page(perPage: 1) {
            media(search: $search, type: ANIME, sort: SEARCH_MATCH) {
              siteUrl
              externalLinks { url site type }
            }
          }
        }`,
        variables: { search: q }
      })
    };
  });

  const responses = UrlFetchApp.fetchAll(requests);

  entries.forEach((e, idx) => {
    const resp = responses[idx];
    let res = null;
    if (resp.getResponseCode() === 200) {
      try {
        const match = JSON.parse(resp.getContentText()).data?.Page?.media?.[0];
        if (match) res = { siteUrl: match.siteUrl, liensExternes: match.externalLinks || [] };
      } catch (err) {}
    }

    if (res) {
      if ((forcer || !e.lien) && res.siteUrl) {
        sheet.getRange(e.ligneAbs, e.colDebut + OFF_LIEN).setValue(res.siteUrl);
        majLien++;
      }
      if (forcer || !e.plateforme) {
        const plat = _plateformeDepuisAniList(res.liensExternes);
        if (plat) {
          sheet.getRange(e.ligneAbs, e.colDebut + OFF_PLATEFORME).setValue(plat);
          majPlat++;
        }
      }
    } else {
      echecs++;
    }
  });

  return { majPlat, majLien, echecs };
}

function _rechercherProchainEpisodeAniList(titre) {
  const requetes = _essaisRechercheAnime(titre);

  for (const q of requetes) {
    try {
      const resp = UrlFetchApp.fetch("https://graphql.anilist.co", {
        method: "post",
        contentType: "application/json",
        muteHttpExceptions: true,
        payload: JSON.stringify({
          query: `query ($search: String) {
            Page(perPage: 5) {
              media(search: $search, type: ANIME, sort: SEARCH_MATCH) {
                siteUrl
                status
                title { romaji english }
                nextAiringEpisode { airingAt episode }
                externalLinks { url site type }
              }
            }
          }`,
          variables: { search: q }
        })
      });

      if (resp.getResponseCode() === 200) {
        const liste = JSON.parse(resp.getContentText()).data?.Page?.media;
        if (liste && liste.length > 0) {
          const match = liste.find(m => m.nextAiringEpisode) || null;
          if (match && match.nextAiringEpisode) {
            return {
              siteUrl: match.siteUrl,
              liensExternes: match.externalLinks || [],
              episode: match.nextAiringEpisode.episode,
              airingAt: match.nextAiringEpisode.airingAt
            };
          }
        }
      }
    } catch (e) {
      Logger.log(`Erreur AniList nextEp (${q}): ` + e.message);
    }
  }
  return null;
}

function _enrichirTMDBCore(apiKey, forcer) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOM_FEUILLE);
  const fullData = sheet.getDataRange().getValues();
  const entries = _parcourirEntrees(fullData, e => _estFilmOuSerie(e.cat) && (forcer || !e.plateforme || !e.lien));
  if (entries.length === 0) return { majPlat: 0, majLien: 0, echecs: 0 };

  let majPlat = 0, majLien = 0, echecs = 0;

  const searchRequests = entries.map(e => {
    const type = e.cat === "📺 Séries" ? "tv" : "movie";
    const q = _essaisRecherche(e.titre)[0] || e.titre;
    const anneeParam = type === "tv" ? `&first_air_date_year=${e.annee}` : `&primary_release_year=${e.annee}`;
    return {
      url: `https://api.themoviedb.org/3/search/${type}?api_key=${apiKey}&language=fr-FR&query=${encodeURIComponent(q)}${anneeParam}`,
      muteHttpExceptions: true
    };
  });

  const searchResponses = UrlFetchApp.fetchAll(searchRequests);
  const foundItems = [];
  const providerRequests = [];

  entries.forEach((e, idx) => {
    const resp = searchResponses[idx];
    let res = null;
    if (resp.getResponseCode() === 200) {
      try {
        const json = JSON.parse(resp.getContentText());
        if (json.results && json.results.length > 0) {
          const type = e.cat === "📺 Séries" ? "tv" : "movie";
          res = { type, id: json.results[0].id, entryIdx: idx };
        }
      } catch (err) {}
    }

    if (res) {
      foundItems.push(res);
      if (forcer || !e.lien) {
        sheet.getRange(e.ligneAbs, e.colDebut + OFF_LIEN).setValue(`https://www.themoviedb.org/${res.type}/${res.id}`);
        majLien++;
      }
      providerRequests.push({
        url: `https://api.themoviedb.org/3/${res.type}/${res.id}/watch/providers?api_key=${apiKey}`,
        muteHttpExceptions: true
      });
    } else {
      echecs++;
    }
  });

  if (providerRequests.length > 0) {
    const providerResponses = UrlFetchApp.fetchAll(providerRequests);
    providerResponses.forEach((pResp, pIdx) => {
      if (pResp.getResponseCode() === 200) {
        try {
          const json = JSON.parse(pResp.getContentText());
          if (json.results && json.results.FR && json.results.FR.flatrate && json.results.FR.flatrate.length > 0) {
            const plat = _mapProviderVersPlateforme(json.results.FR.flatrate[0].provider_name);
            if (plat) {
              const origEntry = entries[foundItems[pIdx].entryIdx];
              sheet.getRange(origEntry.ligneAbs, origEntry.colDebut + OFF_PLATEFORME).setValue(plat);
              majPlat++;
            }
          }
        } catch (err) {}
      }
    });
  }

  return { majPlat, majLien, echecs };
}

function _isoVersJJMMAAAA(iso) {
  const p = (iso || "").toString().slice(0, 10).split("-");
  return (p.length === 3 && p[0].length === 4) ? `${p[2]}/${p[1]}/${p[0]}` : null;
}

function _dateReleaseFR(res, apiKey) {
  if (!res) return null;

  if (res.type === "movie") {
    try {
      const url = `https://api.themoviedb.org/3/movie/${res.id}/release_dates?api_key=${apiKey}`;
      const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const json = JSON.parse(resp.getContentText());
      const fr = (json.results || []).find(r => r.iso_3166_1 === "FR");

      if (fr && fr.release_dates && fr.release_dates.length > 0) {
        const priorites = [3, 2, 4, 5, 6, 1];
        let choisie = null;
        for (const p of priorites) {
          choisie = fr.release_dates.find(d => d.type === p);
          if (choisie) break;
        }
        if (choisie && choisie.release_date) return _isoVersJJMMAAAA(choisie.release_date);
      }
    } catch (e) {
      Logger.log("Erreur _dateReleaseFR (movie): " + e.message);
    }
  }

  return _isoVersJJMMAAAA(res.dateSortie);
}

function _rechercherSortieRecenteTMDB(titre, cat, apiKey) {
  const type = "movie";
  const anneeMin = Math.min(...YEARS);
  const anneeMax = Math.max(...YEARS);

  for (const q of _essaisRecherche(titre)) {
    try {
      const url = `https://api.themoviedb.org/3/search/${type}?api_key=${apiKey}&language=fr-FR&query=${encodeURIComponent(q)}`;
      const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const json = JSON.parse(resp.getContentText());
      if (!json.results || json.results.length === 0) continue;

      const candidats = json.results.slice(0, 8);
      for (const cand of candidats) {
        const dateFR = _dateReleaseFR({ type, id: cand.id, dateSortie: cand.release_date }, apiKey);
        if (dateFR) {
          const annee = parseInt(dateFR.split("/")[2], 10);
          if (annee >= anneeMin && annee <= anneeMax) return { type, id: cand.id, dateFR };
        }
      }
    } catch (e) {
      Logger.log("Erreur _rechercherSortieRecenteTMDB: " + e.message);
    }
  }
  return null;
}

function _rechercherSerieTMDB(titre, apiKey) {
  const anneeMin = Math.min(...YEARS);
  const anneeMax = Math.max(...YEARS);
  const dansPlage = iso => { const a = parseInt((iso || "").slice(0, 4), 10); return a >= anneeMin && a <= anneeMax; };

  for (const q of _essaisRecherche(titre)) {
    try {
      const url = `https://api.themoviedb.org/3/search/tv?api_key=${apiKey}&language=fr-FR&query=${encodeURIComponent(q)}`;
      const json = JSON.parse(UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText());
      if (!json.results || json.results.length === 0) continue;

      for (const cand of json.results.slice(0, 5)) {
        const r = UrlFetchApp.fetch(`https://api.themoviedb.org/3/tv/${cand.id}?api_key=${apiKey}&language=fr-FR`, { muteHttpExceptions: true });
        if (r.getResponseCode() !== 200) continue;
        const tv = JSON.parse(r.getContentText());

        const ne = tv.next_episode_to_air;
        if (ne && ne.air_date && dansPlage(ne.air_date)) {
          return { type: "tv", id: tv.id, dateFR: _isoVersJJMMAAAA(ne.air_date), episode: ne.episode_number, saison: ne.season_number, prochainEp: true };
        }
        if (tv.first_air_date && dansPlage(tv.first_air_date)) {
          return { type: "tv", id: tv.id, dateFR: _isoVersJJMMAAAA(tv.first_air_date), prochainEp: false };
        }
      }
    } catch (e) {
      Logger.log("Erreur _rechercherSerieTMDB: " + e.message);
    }
  }
  return null;
}

function _enrichirIGDBCore(clientId, token, forcer) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOM_FEUILLE);
  const fullData = sheet.getDataRange().getValues();
  const entries = _parcourirEntrees(fullData, e => _estJeu(e.cat) && (forcer || !e.plateforme || !e.lien));
  if (entries.length === 0) return { majPlat: 0, majLien: 0, echecs: 0 };

  let majPlat = 0, majLien = 0, echecs = 0;

  const requests = entries.map(e => {
    const q = _essaisRecherche(e.titre)[0] || e.titre;
    return {
      url: "https://api.igdb.com/v4/games",
      method: "post",
      headers: { "Client-ID": clientId, "Authorization": `Bearer ${token}` },
      payload: `search "${q.replace(/"/g, '\\"')}"; fields name,url,platforms.name,first_release_date; limit 1;`,
      muteHttpExceptions: true
    };
  });

  const responses = UrlFetchApp.fetchAll(requests);

  entries.forEach((e, idx) => {
    const resp = responses[idx];
    let res = null;
    if (resp.getResponseCode() === 200) {
      try {
        const json = JSON.parse(resp.getContentText());
        if (json && json.length > 0) {
          res = {
            siteUrl: json[0].url,
            plateformes: (json[0].platforms || []).map(p => p.name)
          };
        }
      } catch (err) {}
    }

    if (res) {
      if ((forcer || !e.lien) && res.siteUrl) {
        sheet.getRange(e.ligneAbs, e.colDebut + OFF_LIEN).setValue(res.siteUrl);
        majLien++;
      }
      if (forcer || !e.plateforme) {
        const plat = _mapPlateformesJeu(res.plateformes);
        if (plat) {
          sheet.getRange(e.ligneAbs, e.colDebut + OFF_PLATEFORME).setValue(plat);
          majPlat++;
        }
      }
    } else {
      echecs++;
    }
  });

  return { majPlat, majLien, echecs };
}

function _rechercherSortieRecenteIGDB(titre, clientId, token) {
  const anneeMin = Math.min(...YEARS);
  const anneeMax = Math.max(...YEARS);

  for (const q of _essaisRecherche(titre)) {
    try {
      const resp = UrlFetchApp.fetch("https://api.igdb.com/v4/games", {
        method: "post",
        headers: { "Client-ID": clientId, "Authorization": `Bearer ${token}` },
        payload: `search "${q.replace(/"/g, '\\"')}"; fields name,url,platforms.name,first_release_date; limit 10;`,
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() !== 200) continue;

      const json = JSON.parse(resp.getContentText());
      if (!json || json.length === 0) continue;

      for (const cand of json) {
        if (!cand.first_release_date) continue;
        const d = new Date(cand.first_release_date * 1000);
        const annee = d.getUTCFullYear();
        if (annee >= anneeMin && annee <= anneeMax) {
          const dateFR = Utilities.formatDate(d, FUSEAU_FR, "dd/MM/yyyy");
          return {
            siteUrl: cand.url,
            plateformes: (cand.platforms || []).map(p => p.name),
            dateFR
          };
        }
      }
    } catch (e) {
      Logger.log("Erreur _rechercherSortieRecenteIGDB: " + e.message);
    }
  }
  return null;
}

function recupererInfosAnimesAniList() { enrichirAuChoix("1"); }
function recupererInfosTMDB() { enrichirAuChoix("2"); }
function recupererInfosJeuxIGDB() { enrichirAuChoix("3"); }

function _plateformeDepuisAniList(liensExternes) {
  if (!liensExternes || liensExternes.length === 0) return "";
  const streaming = liensExternes.filter(l => l.type === "STREAMING" || /crunchyroll|netflix|adn|wakanim|prime|disney/i.test(l.site || ""));
  if (streaming.length === 0) return "";
  return _mapProviderVersPlateforme(streaming[0].site);
}

// ============================================================
//   WATCHLIST & RECHERCHE AUTOMATIQUE DES SORTIES
// ============================================================

function ajouterALaWatchlist() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const f = _getOrCreateWatchlist(ss);

  const repTitre = ui.prompt("🔭 Ajouter à la watchlist", "Titre exact à surveiller (ex: Dune 3, GTA VI, One Piece...) :", ui.ButtonSet.OK_CANCEL);
  if (repTitre.getSelectedButton() !== ui.Button.OK) return;
  const titre = repTitre.getResponseText().trim();
  if (!titre) return ui.alert("❌ Titre vide.");

  const repType = ui.prompt("Type de contenu", "1 = 🎬 Film\n2 = 📺 Série\n3 = 🎮 Jeu\n4 = 🍿 Anime", ui.ButtonSet.OK_CANCEL);
  if (repType.getSelectedButton() !== ui.Button.OK) return;
  const type = { "1": "film", "2": "serie", "3": "jeu", "4": "anime" }[repType.getResponseText().trim()];
  if (!type) return ui.alert("❌ Choix invalide (1 à 4 attendu).");

  const lastRow = f.getLastRow() + 1;
  f.getRange(lastRow, 1, 1, 2).setValues([[_neutraliserFormule(titre), type]]);
  f.getRange(lastRow, 2, 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(["film", "serie", "jeu", "anime"], true).setAllowInvalid(true).build()
  );

  ui.alert(`✅ "${titre}" (${type}) ajouté à la watchlist.\n\nLance "🤖 Rechercher les nouvelles sorties" quand tu veux vérifier sa date.`);
}

function voirWatchlist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const f = _getOrCreateWatchlist(ss);
  f.showSheet();
  ss.setActiveSheet(f);
}

function synchroniserSortiesAutomatiquement() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOM_FEUILLE);
  if (!sheet) return _alerte(`❌ Onglet "${NOM_FEUILLE}" introuvable.`);

  const fWatch = _getOrCreateWatchlist(ss);
  const data = fWatch.getDataRange().getValues();
  const lignes = data.slice(1);

  if (lignes.length === 0 || lignes.every(r => !r[0])) {
    return _alerte("ℹ️ La watchlist est vide. Utilise \"🔭 Ajouter un titre à surveiller\" d'abord.");
  }

  const props = PropertiesService.getScriptProperties();
  const tmdbKey = props.getProperty(PROP_KEY_TMDB);
  const igdbId = props.getProperty(PROP_KEY_IGDB_CLIENT_ID);
  const igdbSecret = props.getProperty(PROP_KEY_IGDB_CLIENT_SECRET);
  let igdbToken = null;
  if (igdbId && igdbSecret) igdbToken = _obtenirTokenIGDB(igdbId, igdbSecret);

  const fullData = sheet.getDataRange().getValues();
  const existants = new Set(_parcourirEntrees(fullData, () => true).map(e => e.titre.toLowerCase().trim()));

  let ajouts = 0, ignores = 0, echecs = 0;
  const details = [];
  const plage = `${Math.min(...YEARS)}-${Math.max(...YEARS)}`;

  lignes.forEach((row, i) => {
    const titre = (row[0] || "").toString().trim();
    const type = (row[1] || "").toString().trim().toLowerCase();
    if (!titre || !type) return;

    let dateFR = null;
    let cat = CATEGORIES[0];
    let titreFinal = titre;
    let statutMaj = "";

    try {
      if (type === "film") {
        if (!tmdbKey) { statutMaj = "⚠️ Clé TMDB manquante"; }
        else {
          cat = "🎬 Films";
          const res = _rechercherSortieRecenteTMDB(titre, cat, tmdbKey);
          if (res && res.dateFR) {
            dateFR = res.dateFR;
            statutMaj = `✅ Trouvé (${plage}) : ${dateFR}`;
          } else {
            statutMaj = `ℹ️ Aucune sortie ${plage} trouvée sur TMDB pour ce titre`;
          }
        }
      } else if (type === "serie") {
        if (!tmdbKey) { statutMaj = "⚠️ Clé TMDB manquante"; }
        else {
          cat = "📺 Séries";
          const res = _rechercherSerieTMDB(titre, tmdbKey);
          if (res && res.dateFR) {
            dateFR = res.dateFR;
            if (res.prochainEp && res.episode) {
              titreFinal = (res.saison && res.saison > 1)
                ? `${titre} S${res.saison} - Ep. ${res.episode}`
                : `${titre} - Ep. ${res.episode}`;
              statutMaj = `✅ Prochain épisode (${res.episode}) : ${dateFR}`;
            } else {
              statutMaj = `✅ Première diffusion : ${dateFR}`;
            }
          } else {
            statutMaj = `ℹ️ Aucun épisode / lancement ${plage} trouvé sur TMDB`;
          }
        }
      } else if (type === "jeu") {
        if (!igdbToken) { statutMaj = "⚠️ Clés IGDB manquantes/invalides"; }
        else {
          cat = "🎮 Jeux";
          const res = _rechercherSortieRecenteIGDB(titre, igdbId, igdbToken);
          if (res && res.dateFR) {
            dateFR = res.dateFR;
            statutMaj = `✅ Trouvé (${plage}) : ${dateFR}`;
          } else {
            statutMaj = `ℹ️ Aucune sortie ${plage} trouvée sur IGDB pour ce titre`;
          }
        }
      } else if (type === "anime") {
        cat = "🍿 Animés";
        const res = _rechercherProchainEpisodeAniList(titre);
        if (res && res.airingAt) {
          const d = new Date(res.airingAt * 1000);
          dateFR = Utilities.formatDate(d, FUSEAU_FR, "dd/MM/yyyy");
          titreFinal = `${titre} - Ep. ${res.episode}`;
          statutMaj = `✅ Prochain épisode (${res.episode}) : ${dateFR}`;
        } else {
          statutMaj = "ℹ️ Pas de diffusion en cours détectée";
        }
      } else {
        statutMaj = "❌ Type invalide";
      }
    } catch (err) {
      statutMaj = "❌ Erreur : " + err.message;
    }

    fWatch.getRange(i + 2, 3).setValue(Utilities.formatDate(new Date(), FUSEAU_FR, "dd/MM/yyyy HH:mm"));
    fWatch.getRange(i + 2, 4).setValue(statutMaj);

    if (!dateFR) { echecs++; return; }

    if (existants.has(titreFinal.toLowerCase().trim())) {
      ignores++;
      return;
    }

    const msg = insererEntree_(titreFinal, dateFR, cat, "", "", "❌");
    if (msg.startsWith("✅")) {
      ajouts++;
      existants.add(titreFinal.toLowerCase().trim());
      details.push(`• ${titreFinal} → ${dateFR}`);
    } else if (msg.startsWith("⚠️")) {
      ignores++;
    } else {
      echecs++;
    }
  });

  mettreEnValeurProchaineSortie();
  _invaliderCacheJSON();
  SpreadsheetApp.flush();

  let rapport = `🤖 Synchronisation terminée !\n\n✅ ${ajouts} nouvelle(s) sortie(s) ajoutée(s)\n➖ ${ignores} déjà présente(s) / ignorée(s)\n❌ ${echecs} non trouvée(s) ou en erreur`;
  if (details.length > 0) rapport += `\n\nAjouts :\n` + details.join("\n");
  _alerte(rapport);
}

function activerSyncAutoHebdo() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === "synchroniserSortiesAutomatiquement") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("synchroniserSortiesAutomatiquement").timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  SpreadsheetApp.getUi().alert("✅ Recherche automatique activée : tous les lundis à 7h (heure de Paris), le script vérifiera la watchlist et ajoutera les nouvelles sorties trouvées.");
}

function desactiverSyncAutoHebdo() {
  let count = 0;
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === "synchroniserSortiesAutomatiquement") { ScriptApp.deleteTrigger(t); count++; } });
  SpreadsheetApp.getUi().alert(count > 0 ? "🛑 Recherche automatique désactivée." : "ℹ️ Aucune recherche automatique active.");
}

// ============================================================
//   ENRICHISSEMENT MEDIASTACK (PRESSE & ARTICLES)
// ============================================================

function configurerCleMediastack() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const cleActuelle = props.getProperty(PROP_KEY_MEDIASTACK) || "";

  const rep = ui.prompt(
    "🔑 Clé API Mediastack",
    `Colle ta clé API Mediastack (Actuelle : ${cleActuelle ? cleActuelle.slice(0, 6) + "..." : "Aucune"}) :`,
    ui.ButtonSet.OK_CANCEL
  );
  if (rep.getSelectedButton() === ui.Button.OK) {
    const cle = rep.getResponseText().trim();
    if (cle) {
      props.setProperty(PROP_KEY_MEDIASTACK, cle);
      ui.alert("✅ Clé Mediastack enregistrée avec succès !");
    }
  }
}

function enrichirLiensAvecMediastack() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOM_FEUILLE);
  if (!sheet) return ui.alert(`❌ Onglet "${NOM_FEUILLE}" introuvable.`);

  const props = PropertiesService.getScriptProperties();
  let apiKey = props.getProperty(PROP_KEY_MEDIASTACK);

  if (!apiKey) {
    const rep = ui.alert("🔑 Clé manquante", "La clé API Mediastack n'est pas configurée. Veux-tu la renseigner maintenant ?", ui.ButtonSet.YES_NO);
    if (rep === ui.Button.YES) configurerCleMediastack();
    apiKey = props.getProperty(PROP_KEY_MEDIASTACK);
    if (!apiKey) return;
  }

  const repMode = ui.prompt(
    "📰 Recherche d'articles Mediastack",
    "Choisis le mode :\n1 = Uniquement la ligne actuellement sélectionnée\n2 = Toutes les sorties du calendrier sans lien",
    ui.ButtonSet.OK_CANCEL
  );
  if (repMode.getSelectedButton() !== ui.Button.OK) return;

  const mode = repMode.getResponseText().trim();
  let cibles = [];
  const fullData = sheet.getDataRange().getValues();

  if (mode === "1") {
    const activeRange = sheet.getActiveRange();
    if (!activeRange) return ui.alert("❌ Aucune cellule sélectionnée.");
    const row = activeRange.getRow();
    const toutes = _parcourirEntrees(fullData, () => true);
    const trouve = toutes.find(e => e.ligneAbs === row);
    if (!trouve) return ui.alert("❌ Aucune sortie reconnue sur la ligne sélectionnée.");
    cibles.push(trouve);
  } else if (mode === "2") {
    cibles = _parcourirEntrees(fullData, e => !e.lien);
    if (cibles.length === 0) return ui.alert("ℹ️ Toutes les sorties ont déjà un lien.");

    const conf = ui.alert("⚠️ Quota API", `${cibles.length} sortie(s) à traiter.\nContinuer la recherche ?`, ui.ButtonSet.YES_NO);
    if (conf !== ui.Button.YES) return;
  } else {
    return ui.alert("❌ Choix invalide (1 ou 2).");
  }

  let ajouts = 0;
  let nonTrouves = 0;

  const requests = cibles.map(e => {
    const q = _essaisRecherche(e.titre)[0] || e.titre;
    return {
      url: `https://api.apilayer.net/mediastack/news?access_key=${apiKey}&keywords=${encodeURIComponent(q)}&languages=fr&limit=1`,
      headers: { "apikey": apiKey },
      muteHttpExceptions: true
    };
  });

  const responses = UrlFetchApp.fetchAll(requests);

  cibles.forEach((e, idx) => {
    const resp = responses[idx];
    let res = null;
    if (resp.getResponseCode() === 200) {
      try {
        const article = JSON.parse(resp.getContentText()).data?.[0];
        if (article && article.url) {
          res = { url: article.url, source: article.source || "Presse" };
        }
      } catch (err) {}
    }

    if (res) {
      sheet.getRange(e.ligneAbs, e.colDebut + OFF_LIEN).setValue(res.url);
      if (!e.plateforme && res.source) {
        sheet.getRange(e.ligneAbs, e.colDebut + OFF_PLATEFORME).setValue(`📰 ${res.source}`);
      }
      ajouts++;
    } else {
      nonTrouves++;
    }
  });

  _invaliderCacheJSON();
  SpreadsheetApp.flush();

  ui.alert(`✅ Terminé !\n• ${ajouts} article(s) lié(s)\n• ${nonTrouves} non trouvé(s)`);
}

function genererOngletNews() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const key = PropertiesService.getScriptProperties().getProperty(PROP_KEY_MEDIASTACK);
  if (!key) return SpreadsheetApp.getUi().alert("Configure d'abord la clé Mediastack.");

  let fNews = ss.getSheetByName("📰 ACTUALITÉS");
  if (!fNews) {
    fNews = ss.insertSheet("📰 ACTUALITÉS");
    fNews.setTabColor("#ff9900");
  }
  fNews.clear();

  const url = `https://api.apilayer.net/mediastack/news?access_key=${key}&categories=entertainment&languages=fr&limit=15`;
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const data = JSON.parse(response.getContentText()).data || [];

  fNews.getRange(1, 1, 1, 4).setValues([["DATE", "TITRE", "SOURCE", "LIEN"]])
    .setFontWeight("bold").setBackground("#1a1a2e").setFontColor("#fff");

  if (data.length > 0) {
    const rows = data.map(a => [
      (a.published_at || "").slice(0, 10),
      _neutraliserFormule(a.title),
      _neutraliserFormule(a.source),
      _neutraliserFormule(a.url)
    ]);
    fNews.getRange(2, 1, rows.length, 4).setValues(rows);
    fNews.autoResizeColumns(1, 4);
  }
  SpreadsheetApp.getUi().alert(`✅ ${data.length} articles récupérés !`);
}

// ============================================================
//   ENRICHISSEMENT YOUTUBE (BANDES-ANNONCES / CHAÎNES)
// ============================================================

function configurerCleYoutube() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const cleActuelle = props.getProperty(PROP_KEY_YOUTUBE) || "";

  const rep = ui.prompt(
    "🔑 Clé API YouTube",
    `Colle ta clé API YouTube (Actuelle : ${cleActuelle ? cleActuelle.slice(0, 6) + "..." : "Aucune"}) :`,
    ui.ButtonSet.OK_CANCEL
  );
  if (rep.getSelectedButton() === ui.Button.OK) {
    const cle = rep.getResponseText().trim();
    if (cle) {
      props.setProperty(PROP_KEY_YOUTUBE, cle);
      ui.alert("✅ Clé YouTube enregistrée avec succès !");
    }
  }
}

function enrichirLiensAvecYoutube() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOM_FEUILLE);
  if (!sheet) return ui.alert(`❌ Onglet "${NOM_FEUILLE}" introuvable.`);

  const props = PropertiesService.getScriptProperties();
  let apiKey = props.getProperty(PROP_KEY_YOUTUBE);

  if (!apiKey) {
    const rep = ui.alert("🔑 Clé manquante", "La clé API YouTube n'est pas configurée. Veux-tu la renseigner maintenant ?", ui.ButtonSet.YES_NO);
    if (rep === ui.Button.YES) configurerCleYoutube();
    apiKey = props.getProperty(PROP_KEY_YOUTUBE);
    if (!apiKey) return;
  }

  const repMode = ui.prompt(
    "🎥 Recherche de bandes-annonces YouTube",
    "Choisis le mode :\n1 = Uniquement la ligne actuellement sélectionnée\n2 = Toutes les sorties du calendrier sans lien",
    ui.ButtonSet.OK_CANCEL
  );
  if (repMode.getSelectedButton() !== ui.Button.OK) return;

  const mode = repMode.getResponseText().trim();
  let cibles = [];
  const fullData = sheet.getDataRange().getValues();

  if (mode === "1") {
    const activeRange = sheet.getActiveRange();
    if (!activeRange) return ui.alert("❌ Aucune cellule sélectionnée.");
    const row = activeRange.getRow();
    const toutes = _parcourirEntrees(fullData, () => true);
    const trouve = toutes.find(e => e.ligneAbs === row);
    if (!trouve) return ui.alert("❌ Aucune sortie reconnue sur la ligne sélectionnée.");
    cibles.push(trouve);
  } else if (mode === "2") {
    cibles = _parcourirEntrees(fullData, e => !e.lien);
    if (cibles.length === 0) return ui.alert("ℹ️ Toutes les sorties ont déjà un lien.");

    const nbRecherches = cibles.length;
    const msgQuota = nbRecherches > 45
      ? `${cibles.length} sortie(s), mais le quota gratuit YouTube ne permet qu'environ 50 à 100 recherches/jour.\nContinuer quand même ?`
      : `${cibles.length} sortie(s) à traiter.\nContinuer la recherche ?`;
    const conf = ui.alert("⚠️ Quota API", msgQuota, ui.ButtonSet.YES_NO);
    if (conf !== ui.Button.YES) return;
  } else {
    return ui.alert("❌ Choix invalide (1 ou 2).");
  }

  let ajouts = 0;
  let nonTrouves = 0;

  const requests = cibles.map(e => {
    const q = e.titre + " bande annonce trailer";
    return {
      url: "https://www.googleapis.com/youtube/v3/search"
        + "?part=snippet&q=" + encodeURIComponent(q)
        + "&type=video&maxResults=1&order=relevance&relevanceLanguage=fr&safeSearch=moderate&key=" + apiKey,
      muteHttpExceptions: true
    };
  });

  const responses = UrlFetchApp.fetchAll(requests);

  cibles.forEach((e, idx) => {
    const resp = responses[idx];
    let res = null;
    if (resp.getResponseCode() === 200) {
      try {
        const item = JSON.parse(resp.getContentText()).items?.[0];
        if (item && item.id?.videoId) {
          res = "https://www.youtube.com/watch?v=" + item.id.videoId;
        }
      } catch (err) {}
    }

    if (res) {
      sheet.getRange(e.ligneAbs, e.colDebut + OFF_LIEN).setValue(res);
      ajouts++;
    } else {
      nonTrouves++;
    }
  });

  _invaliderCacheJSON();
  SpreadsheetApp.flush();

  ui.alert(`✅ Terminé !\n• ${ajouts} bande(s)-annonce(s) liée(s)\n• ${nonTrouves} non trouvé(s)`);
}

// ============================================================
//   CONFIGURATIONS & FILE D'ATTENTE
// ============================================================

function configurerUrlWebApp() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const rep = ui.prompt("🌐 URL Web App", "Colle l'URL se terminant par /exec :", ui.ButtonSet.OK_CANCEL);
  if (rep.getSelectedButton() === ui.Button.OK) {
    let url = rep.getResponseText().trim().split("?")[0];
    if (url) { props.setProperty("WEBAPP_URL", url); ui.alert("✅ Enregistré !"); }
  }
}

function configurerCleTMDB() {
  const ui = SpreadsheetApp.getUi();
  const rep = ui.prompt("🔑 Clé API TMDB", "Clé :", ui.ButtonSet.OK_CANCEL);
  if (rep.getSelectedButton() === ui.Button.OK) {
    PropertiesService.getScriptProperties().setProperty(PROP_KEY_TMDB, rep.getResponseText().trim());
    ui.alert("✅ Enregistré !");
  }
}

function configurerCleIGDB() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const repId = ui.prompt("🔑 IGDB Client ID", "Client ID :", ui.ButtonSet.OK_CANCEL);
  if (repId.getSelectedButton() !== ui.Button.OK) return;
  const repSecret = ui.prompt("🔑 IGDB Client Secret", "Client Secret :", ui.ButtonSet.OK_CANCEL);
  if (repSecret.getSelectedButton() !== ui.Button.OK) return;
  const id = repId.getResponseText().trim();
  const secret = repSecret.getResponseText().trim();
  if (id && secret) {
    props.setProperty(PROP_KEY_IGDB_CLIENT_ID, id);
    props.setProperty(PROP_KEY_IGDB_CLIENT_SECRET, secret);
    props.deleteProperty(PROP_KEY_IGDB_TOKEN);
    props.deleteProperty(PROP_KEY_IGDB_TOKEN_EXPIRE);
    ui.alert("✅ Enregistré !");
  }
}

function _obtenirTokenIGDB(clientId, clientSecret) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty(PROP_KEY_IGDB_TOKEN);
  const exp = parseInt(props.getProperty(PROP_KEY_IGDB_TOKEN_EXPIRE) || "0", 10);

  if (token && Date.now() < exp) return token;

  const resp = UrlFetchApp.fetch(`https://id.twitch.tv/oauth2/token?client_id=${clientId}&client_secret=${clientSecret}&grant_type=client_credentials`, { method: "post", muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) return null;

  const json = JSON.parse(resp.getContentText());
  props.setProperty(PROP_KEY_IGDB_TOKEN, json.access_token);
  props.setProperty(PROP_KEY_IGDB_TOKEN_EXPIRE, String(Date.now() + (json.expires_in * 1000) - 60000));

  return json.access_token;
}

function configurerCodeAmis() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const rep = ui.prompt("🔑 Code d'accès Amis", "Code secret (vide = aucun code) :", ui.ButtonSet.OK_CANCEL);
  if (rep.getSelectedButton() === ui.Button.OK) {
    const code = rep.getResponseText().trim();
    if (code) props.setProperty(PROP_KEY_CODE_AMIS, code);
    else props.deleteProperty(PROP_KEY_CODE_AMIS);
    ui.alert("✅ Enregistré.");
  }
}

function obtenirLienCalendrierAmis() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  let url = props.getProperty("WEBAPP_URL") || ScriptApp.getService().getUrl();

  if (!url || url.includes("/dev")) {
    configurerUrlWebApp();
    url = props.getProperty("WEBAPP_URL");
  }

  if (!url) return;
  url = url.split("?")[0];
  const code = props.getProperty(PROP_KEY_CODE_AMIS);
  ui.alert("🌐 Liens d'accès", `Formulaire d'ajout :\n${url}?vue=ajouter\n\nCalendrier (WebApp) :\n${url}\n\nAPI JSON :\n${url}?format=json\n\nCode secret actuel : ${code || "Aucun"}`, ui.ButtonSet.OK);
}

function journaliserAjoutPublic_(a, t, d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let f = ss.getSheetByName(NOM_FEUILLE_JOURNAL_AMIS);
  if (!f) {
    f = ss.insertSheet(NOM_FEUILLE_JOURNAL_AMIS);
    f.getRange(1, 1, 1, 4).setValues([["DATE AJOUT", "AUTEUR", "TITRE", "DATE SORTIE"]]).setFontWeight("bold");
    f.hideSheet();
  }
  f.appendRow([Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"), _neutraliserFormule(a), _neutraliserFormule(t), _neutraliserFormule(d)]);
}

function voirJournalAmis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const f = ss.getSheetByName(NOM_FEUILLE_JOURNAL_AMIS);
  if (!f) return SpreadsheetApp.getUi().alert("🧾 Aucun ajout enregistré.");
  f.showSheet();
  ss.setActiveSheet(f);
}

function obtenirLienPageMobile() { obtenirLienCalendrierAmis(); }

function _getOrCreateFeuilleEnAttente(ss) {
  let f = ss.getSheetByName(NOM_FEUILLE_EN_ATTENTE);
  if (!f) {
    f = ss.insertSheet(NOM_FEUILLE_EN_ATTENTE);
    f.setTabColor("#f5b82e");
    f.getRange(1, 1, 1, 9).setValues([["VALIDER ?", "REFUSER ?", "AUTEUR", "TITRE", "DATE SORTIE", "CATÉGORIE", "OÙ VOIR", "LIEN", "PROPOSÉ LE"]])
      .setFontWeight("bold").setBackground("#1e1e34").setFontColor("#ffffff");
    f.setColumnWidths(1, 9, 120);
    f.setFrozenRows(1);
  }
  return f;
}

function _quotaPropositionsAtteint() {
  const cache = CacheService.getScriptCache();
  const cle = "propositions_heure";
  const n = parseInt(cache.get(cle) || "0", 10);
  if (n >= MAX_PROPOSITIONS_PAR_HEURE) return true;
  cache.put(cle, String(n + 1), 3600);
  return false;
}

function ajouterSortiePublique(data) {
  data = data || {};
  const props = PropertiesService.getScriptProperties();
  const codeAttendu = props.getProperty(PROP_KEY_CODE_AMIS);
  if (codeAttendu && (!data.code || data.code.toString().trim() !== codeAttendu.trim())) {
    return "❌ Code d'accès incorrect.";
  }
  if (!data.nom || !data.date) return "❌ Le nom et la date sont obligatoires.";

  const nom = data.nom.toString().trim().slice(0, 120);
  const date = data.date.toString().trim().slice(0, 40);
  if (!nom || !date) return "❌ Le nom et la date sont obligatoires.";

  const parsed = _parseDateEntree(date);
  if (parsed.type === "invalide") return `❌ Format de date non supporté : ${parsed.raison || ""}`;

  if (_quotaPropositionsAtteint()) return "⏳ Trop de propositions en peu de temps, réessaie dans un moment.";

  const categorie = CATEGORIES.includes(data.categorie) ? data.categorie : CATEGORIES[0];
  const lien = _urlSure(data.lien);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const fEnAttente = _getOrCreateFeuilleEnAttente(SpreadsheetApp.getActiveSpreadsheet());
    const lastRow = fEnAttente.getLastRow() + 1;
    fEnAttente.getRange(lastRow, 3, 1, 7).setValues([[
      _neutraliserFormule((data.auteur || "Ami·e").toString().trim().slice(0, 40)),
      _neutraliserFormule(nom),
      _neutraliserFormule(date),
      categorie,
      _neutraliserFormule((data.plateforme || "").toString().trim().slice(0, 60)),
      lien,
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
    ]]);
    fEnAttente.getRange(lastRow, 1, 1, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build()).setValues([[false, false]]);
  } finally {
    lock.releaseLock();
  }

  return "✅ Proposition transmise dans la file d'attente !";
}

function gererEditionValidation(e, sheet) {
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row > 1 && e.value === "TRUE") {
    if (col === 1) _validerLigneEnAttente(sheet, row);
    else if (col === 2) {
      SpreadsheetApp.getActiveSpreadsheet().toast(`🗑️ Refusé : ${sheet.getRange(row, 4).getValue()}`);
      sheet.deleteRow(row);
    }
  }
}

function _validerLigneEnAttente(sheetPending, row) {
  const data = sheetPending.getRange(row, 1, 1, 9).getDisplayValues()[0];
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const msg = insererEntree_(data[3], data[4], data[5], data[6], data[7], "❌");

  if (msg.startsWith("✅")) {
    try { journaliserAjoutPublic_(data[2], data[3], data[4]); } catch (err) { Logger.log("Journal amis : " + err.message); }
    mettreEnValeurProchaineSortie();
    sheetPending.deleteRow(row);
    ss.toast(`✅ Validé : ${data[3]}`);
  } else {
    sheetPending.getRange(row, 1).setValue(false);
    ss.toast(`⚠️ ${msg}`);
  }
}

function voirFeuilleEnAttente() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const f = _getOrCreateFeuilleEnAttente(ss);
  f.showSheet();
  ss.setActiveSheet(f);
}

// ============================================================
//   TRIGGERS & SYNCHRONISATIONS
// ============================================================

function _majQuotidienne() {
  mettreEnValeurProchaineSortie();
}

function activerTriggerQuotidien() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === "_majQuotidienne") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("_majQuotidienne").timeBased().everyDays(1).atHour(6).create();
  SpreadsheetApp.getUi().alert("✅ MAJ quotidienne activée.");
}

function desactiverTriggerQuotidien() {
  let count = 0;
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === "_majQuotidienne") { ScriptApp.deleteTrigger(t); count++; } });
  SpreadsheetApp.getUi().alert(count > 0 ? `🛑 MAJ quotidienne désactivée.` : "ℹ️ Aucune MAJ active.");
}

function activerRappelEmail() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === "envoyerRappelsSorties") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("envoyerRappelsSorties").timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(9).create();
  SpreadsheetApp.getUi().alert("✅ Rappels email activés (Dimanche 9h).");
}

function desactiverRappelEmail() {
  let count = 0;
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === "envoyerRappelsSorties") { ScriptApp.deleteTrigger(t); count++; } });
  SpreadsheetApp.getUi().alert(count > 0 ? `🛑 Rappels email désactivés.` : "ℹ️ Aucun rappel actif.");
}

function envoyerRappelsSorties() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOM_FEUILLE);
  const em = Session.getEffectiveUser().getEmail() || Session.getActiveUser().getEmail();
  if (!em || !sheet) return;

  const fullData = sheet.getDataRange().getValues();
  const t = new Date(); t.setHours(0,0,0,0);
  const aRap = _parcourirEntrees(fullData, e => !e.indetermine && e.jourActuel !== "?" && e.jourActuel !== "" && e.statut !== "✔️").filter(e => {
    const d = new Date(e.annee, e.mois, parseInt(e.jourActuel, 10));
    d.setHours(0,0,0,0);
    const diff = Math.round((d - t) / 86400000);
    return diff >= 0 && diff <= 7;
  });

  if (aRap.length > 0) {
    MailApp.sendEmail(em, `📅 Rappel : ${aRap.length} sorties prévues dans les 7 jours`, aRap.map(e => `• ${e.titre} - ${e.jourActuel}/${e.mois+1}/${e.annee} (${e.cat})`).join("\n"));
  }
}

function retrierTousLesMois() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOM_FEUILLE);
  if (!sheet) return;

  const fullData = sheet.getDataRange().getValues();

  YEARS.forEach(annee => {
    for (let m = 0; m < 12; m++) {
      const info = _infoMois(fullData, annee, m);
      if (info.ligneData === -1) continue;

      const vals = sheet.getRange(info.ligneData, info.colDebut, info.nbLignesData, COLS_PAR_MOIS).getValues();
      const remplis = vals.filter(r => (r[OFF_TITRE] || "").toString().trim() !== "");
      if (remplis.length < 2) continue;

      remplis.sort((a, b) => {
        const ja = (!a[OFF_JOUR] || a[OFF_JOUR] === "?") ? 99 : (parseInt(a[OFF_JOUR], 10) || 99);
        const jb = (!b[OFF_JOUR] || b[OFF_JOUR] === "?") ? 99 : (parseInt(b[OFF_JOUR], 10) || 99);
        return ja - jb;
      });

      const vides = Array(info.nbLignesData - remplis.length).fill(0).map(() => Array(COLS_PAR_MOIS).fill(""));
      sheet.getRange(info.ligneData, info.colDebut, info.nbLignesData, COLS_PAR_MOIS).setValues([...remplis, ...vides]);
      _stylerBlocMois(sheet, info, m, [...remplis, ...vides]);
    }

    const infoI = _infoIndetermine(fullData, annee);
    if (infoI.ligneData !== -1) {
      const valsI = sheet.getRange(infoI.ligneData, 1, infoI.nbLignesData, COLS_PAR_MOIS).getValues();
      const remplisI = valsI.filter(r => (r[OFF_TITRE] || "").toString().trim() !== "");
      if (remplisI.length >= 2) {
        remplisI.sort((a, b) => _cleSortPeriode(a[OFF_JOUR]) - _cleSortPeriode(b[OFF_JOUR]));
        const videsI = Array(infoI.nbLignesData - remplisI.length).fill(0).map(() => Array(COLS_PAR_MOIS).fill(""));
        sheet.getRange(infoI.ligneData, 1, infoI.nbLignesData, COLS_PAR_MOIS).setValues([...remplisI, ...videsI]);
        _stylerBlocIndetermine(sheet, infoI, [...remplisI, ...videsI]);
      }
    }
  });

  mettreEnValeurProchaineSortie();
  _invaliderCacheJSON();
  SpreadsheetApp.getUi().alert("✅ Tous les mois et indéterminés ont été triés chronologiquement !");
}

function _getOrCreateDashboard(ss) {
  let f = ss.getSheetByName(NOM_FEUILLE_DASHBOARD);
  if (!f) {
    f = ss.insertSheet(NOM_FEUILLE_DASHBOARD);
    f.setTabColor("#2d6a4f");
  }
  return f;
}

function actualiserTableauDeBord() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOM_FEUILLE);
  let f = _getOrCreateDashboard(ss);
  f.clear();

  const fullData = sheet.getDataRange().getValues();
  const aVoir = _parcourirEntrees(fullData, e => e.statut !== "✔️");
  f.getRange(1, 1, 1, 5).merge().setValue("📋 TABLEAU DE BORD (En attente de visionnage)")
    .setBackground("#1a1a2e").setFontColor("#fff").setFontWeight("bold").setHorizontalAlignment("center");

  f.getRange(2, 1, 1, 5).setValues([["TITRE", "DATE DE SORTIE", "CATÉGORIE", "OÙ VOIR", "STATUT"]])
    .setBackground("#313b4c").setFontColor("#ffffff").setFontWeight("bold");

  if (aVoir.length > 0) {
    f.getRange(3, 1, aVoir.length, 5).setValues(aVoir.map(e => [
      _neutraliserFormule(e.titre),
      e.indetermine ? `${e.periodeActuelle} ${e.annee}` : `${(e.jourActuel && e.jourActuel !== "?") ? e.jourActuel : "?"}/${e.mois+1}/${e.annee}`,
      e.cat,
      _neutraliserFormule(e.plateforme || "—"),
      e.statut || "❌"
    ]));
    f.autoResizeColumns(1, 5);
  }
  SpreadsheetApp.getUi().alert(`✅ Tableau de bord actualisé (${aVoir.length} œuvres à voir).`);
}

function synchroniserGoogleCalendar() {
  let count = 0;
  let cal = CalendarApp.getCalendarsByName(NOM_CALENDRIER_GOOGLE)[0] || CalendarApp.createCalendar(NOM_CALENDRIER_GOOGLE);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOM_FEUILLE);
  const fullData = sheet.getDataRange().getValues();

  _parcourirEntrees(fullData, e => !e.indetermine && e.jourActuel !== "?" && e.jourActuel !== "").forEach(e => {
    const d = new Date(e.annee, e.mois, parseInt(e.jourActuel, 10));
    if (!cal.getEventsForDay(d).some(ev => ev.getTitle() === e.titre)) {
      cal.createAllDayEvent(e.titre, d, { description: e.plateforme });
      count++;
    }
  });
  SpreadsheetApp.getUi().alert(`📆 Synchronisé : ${count} événements ajoutés.`);
}

function enregistrerAjoutRecent_(e) {
  const p = PropertiesService.getUserProperties();
  let l = [];
  try { l = JSON.parse(p.getProperty(PROP_KEY_RECENTS) || "[]"); } catch(ex){}
  l.unshift({ id: Utilities.getUuid(), titre: e.titre, ts: Date.now() });
  p.setProperty(PROP_KEY_RECENTS, JSON.stringify(l.slice(0, NB_RECENTS_MAX)));
}

function obtenirDernieresSorties() {
  _exigerContexteTableur();
  try { return JSON.parse(PropertiesService.getUserProperties().getProperty(PROP_KEY_RECENTS) || "[]"); } catch (e) { return []; }
}

function annulerAjoutDepuisSidebar(id) {
  _exigerContexteTableur();
  const p = PropertiesService.getUserProperties();
  let l = JSON.parse(p.getProperty(PROP_KEY_RECENTS) || "[]");
  const it = l.find(x => x.id === id);
  if (!it) return { erreur: "Élément introuvable ou trop ancien." };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOM_FEUILLE);
  const res = _parcourirEntrees(sheet, e => e.titre.toLowerCase().includes(it.titre.toLowerCase()));
  if (res.length > 0) {
    supprimerEntreeInterne_(ss, res[0], true);
    mettreEnValeurProchaineSortie();
  }
  p.setProperty(PROP_KEY_RECENTS, JSON.stringify(l.filter(x => x.id !== id)));
  return { ok: true, titre: it.titre };
}

function ajouterSortiesDepuisHTMLSimple(texte) {
  _exigerContexteTableur();
  let ajoutees = 0;
  const erreurs = [];
  texte.split("\n").filter(l => l.trim() !== "").forEach(l => {
    const p = l.split("|").map(x => x.trim());
    if (p[0] && p[1]) {
      const msg = insererEntree_(p[0], p[1], p[2] || CATEGORIES[0], p[3] || "", p[4] || "", p[5] || "❌");
      if (msg.startsWith("✅")) ajoutees++;
      else erreurs.push(`${p[0]} : ${msg}`);
    }
  });
  mettreEnValeurProchaineSortie();
  return { ajoutees, erreurs };
}

// ============================================================
//   API JSON PUBLIQUE  (?format=json)
// ============================================================

function configurerTokenAPI() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const actuel = props.getProperty(PROP_KEY_API_TOKEN);
  const rep = ui.prompt(
    "🔐 Token de l'API JSON",
    `Token actuel : ${actuel ? "défini (" + actuel.slice(0, 3) + "…)" : "aucun (API ouverte à tous)"}\n\nSaisis un mot de passe (ex : une longue suite de caractères) ou laisse vide pour rendre l'API publique :`,
    ui.ButtonSet.OK_CANCEL
  );
  if (rep.getSelectedButton() !== ui.Button.OK) return;
  const t = rep.getResponseText().trim();
  if (t) props.setProperty(PROP_KEY_API_TOKEN, t);
  else props.deleteProperty(PROP_KEY_API_TOKEN);
  ui.alert(t ? "✅ Token enregistré. Ajoute &token=TON_TOKEN à tes appels." : "✅ Token supprimé : l'API est publique.");
}

function testerAPI() {
  const items = _construireItemsAPI();
  const avecDate = items.filter(i => i.isoDate).length;
  const url = PropertiesService.getScriptProperties().getProperty("WEBAPP_URL") || ScriptApp.getService().getUrl() || "(URL non configurée)";
  _alerte(`🧪 L'API lit ${items.length} sorties (${avecDate} avec jour précis).\n\nExemple :\n${JSON.stringify(items[0] || {}, null, 2)}\n\nURL : ${url}?format=json`);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.format === "json") return _apiJSON(p);
  if (p.format === "tmdb") return _apiTMDB(p);
  if (p.vue === "ajouter") return _pageAjoutPublic();
  return _pageConsultationPublic();
}

// Construit la liste complète des sorties (avec cache)
function _construireItemsAPI() {
  const cache = _lireCacheDecoupe(CACHE_KEY_JSON);
  if (cache) { try { return JSON.parse(cache); } catch (e) {} }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOM_FEUILLE);
  const fullData = sheet.getDataRange().getValues();

  const items = _parcourirEntrees(fullData, () => true).map(e => {
    const jour = (!e.indetermine && e.jourActuel && e.jourActuel !== "?") ? parseInt(e.jourActuel, 10) : null;
    const mois = e.indetermine ? null : e.mois + 1;
    return {
      titre: e.titre,
      annee: e.annee,
      mois: mois,
      jour: isNaN(jour) ? null : jour,
      periode: e.indetermine ? (e.periodeActuelle || String(e.annee)) : null,
      isoDate: (jour && !isNaN(jour)) ? `${e.annee}-${_pad2(mois)}-${_pad2(jour)}` : null,
      categorie: e.cat,
      plateforme: e.plateforme,
      vu: _statutEstVu_(e.statut),
      lien: _urlSure(e.lien)
    };
  });

  _sauvegarderCacheDecoupe(JSON.stringify(items), CACHE_KEY_JSON);
  return items;
}

function _sortieJSON(obj, callback) {
  const txt = JSON.stringify(obj);
  if (callback && /^[a-zA-Z0-9_]+$/.test(callback)) {
    return ContentService.createTextOutput(`${callback}(${txt});`).setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON);
}

function _dateModifTableurISO_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const fichier = DriveApp.getFileById(ss.getId());
    const date = fichier.getLastUpdated();
    return date && !isNaN(date.getTime()) ? date.toISOString() : new Date().toISOString();
  } catch (e) {
    return new Date().toISOString();
  }
}

function _apiJSON(p) {
  // Protection optionnelle par token
  const tokenAttendu = PropertiesService.getScriptProperties().getProperty(PROP_KEY_API_TOKEN);
  if (tokenAttendu && p.token !== tokenAttendu) {
    return _sortieJSON({ error: "Token manquant ou invalide." }, p.callback);
  }

  let items;
  try { items = _construireItemsAPI(); }
  catch (err) { return _sortieJSON({ error: "Erreur serveur : " + err.message }, p.callback); }

  // Filtres optionnels
  if (p.annee)     items = items.filter(i => i.annee === parseInt(p.annee, 10));
  if (p.mois)      items = items.filter(i => i.mois === parseInt(p.mois, 10));
  if (p.categorie) items = items.filter(i => _normSimple(i.categorie).includes(_normSimple(p.categorie)));
  if (p.plateforme) items = items.filter(i => _normSimple(i.plateforme).includes(_normSimple(p.plateforme)));
  if (p.q)         items = items.filter(i => _normSimple(i.titre).includes(_normSimple(p.q)));
  if (p.vu)        items = items.filter(i => i.vu === (p.vu === "true"));
  if (p.from)      items = items.filter(i => i.isoDate && i.isoDate >= p.from);
  if (p.to)        items = items.filter(i => i.isoDate && i.isoDate <= p.to);

  // Tri chronologique (dates inconnues à la fin)
  items = items.slice().sort((a, b) => (a.isoDate || "9999").localeCompare(b.isoDate || "9999"));

  const total = items.length;
  const offset = Math.max(0, parseInt(p.offset, 10) || 0);
  if (offset) items = items.slice(offset);
  if (p.limit) items = items.slice(0, Math.max(1, parseInt(p.limit, 10) || 1));

  return _sortieJSON({ count: items.length, total: total, updated: new Date().toISOString(), sourceUpdated: _dateModifTableurISO_(), offset: offset, items: items }, p.callback);
}

// Proxy TMDB restreint aux recherches et fiches nécessaires au calendrier.
// La clé est lue dans les propriétés du script et n'est jamais envoyée au navigateur.
function _apiTMDB(p) {
  const apiKey = PropertiesService.getScriptProperties().getProperty(PROP_KEY_TMDB);
  if (!apiKey) return _sortieJSON({ error: "Clé TMDB non configurée dans Apps Script." }, p.callback);

  const mode = (p.mode || "").toString();
  const query = (p.q || "").toString().trim().slice(0, 120);
  const id = (p.id || "").toString().trim();
  const media = p.media === "tv" ? "tv" : "movie";
  let path = "";

  if (mode === "multi" || mode === "movie") {
    if (!query) return _sortieJSON({ error: "Recherche vide." }, p.callback);
    path = mode === "multi" ? "search/multi" : "search/movie";
    path += "?query=" + encodeURIComponent(query) + "&language=fr-FR&include_adult=false";
    if (mode === "movie") path += "&region=FR";
  } else if (mode === "details" || mode === "videos") {
    if (!/^\d{1,10}$/.test(id)) return _sortieJSON({ error: "Identifiant TMDB invalide." }, p.callback);
    path = media + "/" + id + (mode === "videos" ? "/videos" : "") + "?language=fr-FR";
  } else if (mode === "releases") {
    if (!/^\d{1,10}$/.test(id)) return _sortieJSON({ error: "Identifiant TMDB invalide." }, p.callback);
    path = "movie/" + id + "/release_dates";
  } else {
    return _sortieJSON({ error: "Mode TMDB non autorisé." }, p.callback);
  }

  const cacheInput = mode + "|" + media + "|" + query + "|" + id;
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, cacheInput, Utilities.Charset.UTF_8);
  const suffix = Utilities.base64EncodeWebSafe(digest).replace(/=+$/g, "").slice(0, 32);
  const cacheKey = "tmdb_calendar_" + suffix;
  const cache = CacheService.getScriptCache();
  try {
    const cached = cache.get(cacheKey);
    if (cached) return _sortieJSON(JSON.parse(cached), p.callback);
    const url = "https://api.themoviedb.org/3/" + path + (path.indexOf("?") >= 0 ? "&" : "?") + "api_key=" + encodeURIComponent(apiKey);
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const status = response.getResponseCode();
    if (status < 200 || status >= 300) return _sortieJSON({ error: "TMDB indisponible (" + status + ")." }, p.callback);
    const body = response.getContentText("UTF-8");
    const payload = JSON.parse(body);
    if (body.length < 60000) {
      try { cache.put(cacheKey, body, mode === "multi" || mode === "movie" ? 21600 : 86400); } catch (cacheError) {}
    }
    return _sortieJSON(payload, p.callback);
  } catch (err) {
    Logger.log("Erreur proxy TMDB : " + err.message);
    return _sortieJSON({ error: "Impossible de récupérer les informations TMDB." }, p.callback);
  }
}

function _pageAjoutPublic() {
  const codeRequis = !!PropertiesService.getScriptProperties().getProperty(PROP_KEY_CODE_AMIS);

  const html = `<!DOCTYPE html>
  <html lang="fr">
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Proposer une sortie — L'Œil Critique</title>
    <link href="https://fonts.googleapis.com/css2?family=Bebas+Neue&family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    <style>
      :root {
        --bg-gradient: radial-gradient(130% 120% at 50% 0%, #1e1e34 0%, #0c0c14 70%);
        --gold: #F5B82E;
        --gold-dark: #C68E13;
        --ticket-bg: #FAF7F2;
        --ink: #141523;
        --ink-soft: #5C5F79;
        --line: #E2DDD0;
        --ok: #2D6A4F;
        --err: #E63946;
      }
      * { box-sizing: border-box; }
      body {
        margin: 0; padding: 24px 16px 50px;
        background: var(--bg-gradient);
        font-family: 'Inter', -apple-system, sans-serif;
        color: #fff; min-height: 100vh; max-width: 540px; margin: auto;
      }
      .header { text-align: center; margin-bottom: 20px; }
      .header .kicker { font-size: 10px; letter-spacing: 3px; color: var(--gold); font-weight: 800; text-transform: uppercase; }
      .header h1 { font-family: 'Bebas Neue', sans-serif; font-size: 36px; color: #FFFFFF; margin: 4px 0 0; line-height: 1; letter-spacing: 1px; }
      .form-ticket {
        background: var(--ticket-bg); color: var(--ink);
        border-radius: 18px; padding: 24px;
        box-shadow: 0 16px 40px rgba(0,0,0,0.6);
        border: 1px solid rgba(245, 184, 46, 0.2);
      }
      .form-header {
        font-family: 'Bebas Neue', sans-serif; font-size: 22px; color: var(--ink);
        letter-spacing: 1px; margin-bottom: 16px; display: flex; align-items: center;
        justify-content: space-between; border-bottom: 2px dashed rgba(20, 21, 35, 0.15); padding-bottom: 10px;
      }
      .field { margin-bottom: 14px; }
      label { font-size: 10.5px; font-weight: 800; text-transform: uppercase; letter-spacing: .6px; color: var(--ink-soft); display: block; margin-bottom: 5px; }
      input, select {
        width: 100%; padding: 11px 13px; border-radius: 10px; border: 1.5px solid var(--line);
        font-size: 14px; font-family: inherit; background: #fff; color: var(--ink); font-weight: 600;
      }
      input:focus, select:focus { outline: none; border-color: var(--gold-dark); box-shadow: 0 0 0 3px rgba(245,184,46,0.25); }
      button.btn-submit {
        width: 100%; padding: 14px; background: linear-gradient(180deg, var(--gold) 0%, var(--gold-dark) 100%);
        border: none; border-radius: 999px; font-weight: 800; font-size: 14px; color: var(--ink);
        cursor: pointer; margin-top: 12px; box-shadow: 0 6px 20px rgba(245, 184, 46, 0.35);
      }
      button.btn-submit:disabled { opacity: 0.6; cursor: not-allowed; }
      .msg { margin-top: 14px; padding: 12px; border-radius: 10px; font-size: 13px; display: none; line-height: 1.45; text-align: center; }
      .msg.ok { background: #DFF1E7; color: var(--ok); border: 1px solid #bce2cb; display: block; }
      .msg.err { background: #FDECEC; color: var(--err); border: 1px solid #F6C4C4; display: block; }
    </style>
  </head>
  <body>
    <div class="header">
      <div class="kicker">L'Œil Critique</div>
      <h1>Proposer une sortie</h1>
    </div>
    <div class="form-ticket">
      <div class="form-header">
        <span>➕ Nouvelle proposition</span>
        <span style="font-size:12px; font-family:sans-serif; color:var(--ink-soft); font-weight:normal;">File d'attente</span>
      </div>
      ${codeRequis ? `<div class="field"><label>Code d'accès</label><input type="text" id="p-code" placeholder="Code secret"></div>` : ""}
      <div class="field"><label>Ton prénom / pseudo</label><input type="text" id="p-auteur" placeholder="Ex: Alex" maxlength="40"></div>
      <div class="field"><label>Nom de l'œuvre</label><input type="text" id="p-nom" placeholder="Ex: Dune 3, GTA VI..." maxlength="120"></div>
      <div class="field">
        <label>Rythme de diffusion</label>
        <select id="p-rythme">
          <option value="">Unique (Film, Saison complète...)</option>
          <option value="[Hebdo]">Hebdomadaire</option>
          <option value="[Quotidien]">Quotidien</option>
        </select>
      </div>
      <div class="field"><label>Date de sortie</label><input type="text" id="p-date" placeholder="Ex: 15/10/2026, Automne 2026..." maxlength="40"></div>
      <div class="field"><label>Catégorie</label><select id="p-cat">${CATEGORIES.map(c => `<option value="${c}">${c}</option>`).join("")}</select></div>
      <div class="field"><label>Où voir / Plateforme</label><input type="text" id="p-plat" placeholder="Ex: 📺 Netflix, 🎥 Cinéma..." maxlength="60"></div>
      <button class="btn-submit" id="btnSubmit" onclick="envoyer()">Envoyer la proposition</button>
      <div id="msg" class="msg" role="status" aria-live="polite"></div>
    </div>
    <script>
      function envoyer() {
        const msg = document.getElementById("msg");
        const btn = document.getElementById("btnSubmit");
        let nom = document.getElementById("p-nom").value.trim();
        const date = document.getElementById("p-date").value.trim();
        const rythme = document.getElementById("p-rythme").value;

        if (!nom || !date) {
          msg.className = "msg err";
          msg.textContent = "❌ Le nom et la date sont obligatoires.";
          msg.style.display = "block";
          return;
        }

        if (rythme) nom = nom + " " + rythme;

        btn.disabled = true;
        btn.textContent = "Transmission…";
        msg.style.display = "none";

        const data = {
          code: document.getElementById("p-code") ? document.getElementById("p-code").value : "",
          auteur: document.getElementById("p-auteur").value,
          nom: nom,
          date: date,
          categorie: document.getElementById("p-cat").value,
          plateforme: document.getElementById("p-plat").value
        };

        google.script.run
          .withSuccessHandler(r => {
            btn.disabled = false;
            btn.textContent = "Envoyer la proposition";
            const isOk = r && r.startsWith("✅");
            msg.className = "msg " + (isOk ? "ok" : "err");
            msg.textContent = r;
            msg.style.display = "block";
            if (isOk) {
              document.getElementById("p-nom").value = "";
              document.getElementById("p-date").value = "";
              document.getElementById("p-plat").value = "";
            }
          })
          .withFailureHandler(err => {
            btn.disabled = false;
            btn.textContent = "Envoyer la proposition";
            msg.className = "msg err";
            msg.textContent = "❌ " + err.message;
            msg.style.display = "block";
          })
          .ajouterSortiePublique(data);
      }
    </script>
  </body>
  </html>`;
  return HtmlService.createHtmlOutput(html);
}

function _pageConsultationPublic() {
  let jsonCache = _lireCacheDecoupe(CACHE_KEY_PAGE);
  let itemsData = [];

  if (jsonCache) {
    try { itemsData = JSON.parse(jsonCache); } catch(e) {}
  }

  if (!itemsData || itemsData.length === 0) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOM_FEUILLE);
    const fullData = sheet.getDataRange().getValues();
    itemsData = _parcourirEntrees(fullData, () => true).map(e => {
      let dateStr = "";
      let isoDate = null;
      let ts = 9999999999999;
      if (e.indetermine) {
        dateStr = `${e.periodeActuelle} ${e.annee}`;
      } else {
        const jStr = (e.jourActuel && e.jourActuel !== "?") ? String(e.jourActuel).padStart(2, "0") : null;
        const mStr = String((e.mois || 0) + 1).padStart(2, "0");
        dateStr = `${jStr ? jStr + "/" : ""}${mStr}/${e.annee}`;
        if (jStr) {
          isoDate = `${e.annee}-${mStr}-${jStr}`;
          ts = new Date(e.annee, e.mois, parseInt(e.jourActuel, 10)).getTime();
        }
      }
      return {
        titre: e.titre,
        dateAff: dateStr,
        isoDate: isoDate,
        cat: e.cat,
        plateforme: e.plateforme,
        vu: e.statut || "❌",
        ts: ts,
        lien: _urlSure(e.lien)
      };
    });
    _sauvegarderCacheDecoupe(JSON.stringify(itemsData), CACHE_KEY_PAGE);
  }

  const payloadSafe = JSON.stringify(itemsData).replace(/<\/script>/gi, "<\\/script>").replace(/<!--/g, "<\\!--");

  const html = `<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>L'Œil Critique</title><link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap" rel="stylesheet"><style>
    :root { --bg: #0f0f18; --surface: #1e1e34; --gold: #F5B82E; --text: #fff; --text-muted: #8b8b9e; }
    body { font-family: 'Inter', sans-serif; background: var(--bg); color: var(--text); margin: 0; padding: 20px 14px; }
    .container { max-width: 600px; margin: auto; }
    h1 { text-align: center; font-size: 24px; margin-bottom: 20px; font-weight: 800; }
    .search-box { width: 100%; padding: 14px 18px; border-radius: 12px; background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.1); color: #fff; font-size: 15px; margin-bottom: 20px; outline: none; box-sizing: border-box;}
    .tabs { display: flex; background: rgba(255,255,255,0.05); border-radius: 10px; padding: 4px; margin-bottom: 16px; gap: 4px; }
    .tab { flex: 1; text-align: center; padding: 10px 4px; border-radius: 8px; font-size: 12px; font-weight: 700; cursor: pointer; color: var(--text-muted); }
    .tab.active { background: var(--surface); color: var(--text); }
    .filters { display: flex; flex-wrap: nowrap; overflow-x: auto; gap: 8px; margin-bottom: 24px; padding-bottom: 8px; -webkit-overflow-scrolling: touch; scrollbar-width: none; }
    .filters::-webkit-scrollbar { display: none; }
    .filter-btn { flex-shrink: 0; background: var(--surface); border: 1px solid transparent; color: var(--text-muted); padding: 8px 16px; border-radius: 20px; font-size: 12px; font-weight: 700; cursor: pointer; }
    .filter-btn.active { background: rgba(245, 184, 46, 0.15); color: var(--gold); border-color: var(--gold); }
    .card { background: var(--surface); padding: 16px; border-radius: 14px; margin-bottom: 12px; display: flex; justify-content: space-between; align-items: center; border-left: 4px solid var(--surface); }
    .card-content { flex: 1; }
    .title { font-size: 15px; font-weight: 800; display: flex; align-items: center; flex-wrap: wrap; gap: 6px; }
    .meta { font-size: 12.5px; color: var(--text-muted); margin-top: 6px; font-weight: 600; }
    .status-badge { font-size: 10.5px; font-weight: 800; text-transform: uppercase; padding: 4px 8px; border-radius: 6px; margin-left: 10px; }
    .bg-today { background: #2e7d32; color: #ffffff; } .bg-soon { background: rgba(45, 106, 79, 0.2); color: #4ade80; border: 1px solid rgba(74, 222, 128, 0.3); } .bg-past { background: rgba(255, 255, 255, 0.05); color: #8b8b9e; }
    .tag-hebdo { background: #4f46e5; color: #fff; font-size: 9px; padding: 2px 6px; border-radius: 4px; font-weight: 800; text-transform: uppercase; }
    .empty { text-align: center; padding: 40px 20px; color: var(--text-muted); font-size: 14px; font-weight: 600; }
  </style></head><body>
  <div class="container"><h1>🎬 L'Œil Critique</h1>
    <input type="text" id="searchInput" class="search-box" placeholder="Rechercher un titre...">
    <div class="tabs" id="timeTabs">
      <div class="tab" data-time="all">Toutes</div>
      <div class="tab" data-time="today">🔥 Aujourd'hui</div>
      <div class="tab active" data-time="future">À venir</div>
      <div class="tab" data-time="past">Sortis</div>
    </div>
    <div class="filters" id="catFilters"><button class="filter-btn active" data-cat="all">Tout</button>
      ${CATEGORIES.map(c => `<button class="filter-btn" data-cat="${c}">${c.split(' ')[1] || c}</button>`).join("")}
    </div><div id="results"></div>
  </div>
  <script>
    const items = ${payloadSafe};
    const now = new Date();
    const pad = n => String(n).padStart(2, '0');
    const todayStr = \`\${now.getFullYear()}-\${pad(now.getMonth()+1)}-\${pad(now.getDate())}\`;
    let state = { search: "", cat: "all", time: "future" };

    document.getElementById("searchInput").addEventListener("input", e => { state.search = e.target.value.toLowerCase(); render(); });

    document.getElementById("timeTabs").addEventListener("click", e => {
      if(!e.target.classList.contains("tab")) return;
      document.querySelectorAll(".tab").forEach(t => t.classList.remove("active"));
      e.target.classList.add("active");
      state.time = e.target.getAttribute("data-time");
      render();
    });

    document.getElementById("catFilters").addEventListener("click", e => {
      if(!e.target.classList.contains("filter-btn")) return;
      document.querySelectorAll(".filter-btn").forEach(t => t.classList.remove("active"));
      e.target.classList.add("active");
      state.cat = e.target.getAttribute("data-cat");
      render();
    });

    function render() {
      const container = document.getElementById("results");
      container.innerHTML = "";

      const filtered = items.filter(item => {
        if(state.search && !item.titre.toLowerCase().includes(state.search)) return false;
        if(state.cat !== "all" && item.cat !== state.cat) return false;

        if(state.time === "today") {
          return item.isoDate === todayStr;
        } else if(state.time === "future") {
          if (item.isoDate && item.isoDate < todayStr) return false;
        } else if(state.time === "past") {
          if (!item.isoDate || item.isoDate >= todayStr) return false;
        }
        return true;
      });

      if(filtered.length === 0) return container.innerHTML = '<div class="empty">Aucun résultat. 🕵️‍♂️</div>';

      filtered.sort((a,b) => a.ts - b.ts);
      let html = "";

      filtered.forEach(item => {
        let badge = "";
        if (item.isoDate) {
          if (item.isoDate === todayStr) {
            badge = '<div class="status-badge bg-today">Aujourd\\'hui</div>';
          } else if (item.isoDate < todayStr) {
            badge = '<div class="status-badge bg-past">Sorti</div>';
          } else if (item.ts !== 9999999999999) {
            const diff = Math.round((new Date(item.isoDate + "T00:00:00") - new Date(todayStr + "T00:00:00")) / 86400000);
            if (diff <= 30) badge = \`<div class="status-badge bg-soon">J-\${diff}</div>\`;
          }
        }

        let titreHTML = escapeHTML(item.titre).replace(/\\[Hebdo(.*?)\\]/g, '<span class="tag-hebdo">Hebdo$1</span>').replace(/\\[Quotidien\\]/g, '<span class="tag-hebdo" style="background:#0ea5e9;">Quotidien</span>');
        const safeLien = /^https?:\\/\\//i.test(item.lien || "") ? escapeHTML(item.lien) : "";

        html += \`<div class="card"><div class="card-content"><div class="title">\${titreHTML} \${item.vu ? '<span style="font-size:11px;">'+escapeHTML(item.vu)+'</span>':''} \${safeLien ? ' <a href="'+safeLien+'" target="_blank" rel="noopener noreferrer" style="text-decoration:none;">🔗</a>':''}</div><div class="meta">\${escapeHTML(item.cat)} • <span style="color:#fff;">\${escapeHTML(item.dateAff)}</span> \${item.plateforme ? '• ' + escapeHTML(item.plateforme) : ''}</div></div>\${badge}</div>\`;
      });
      container.innerHTML = html;
    }

    function escapeHTML(s) { return (s||"").toString().replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#039;"); }
    render();
  </script></body></html>`;

  return HtmlService.createHtmlOutput(html);
}

// ============================================================
//   EXPLORATEUR LATÉRAL
// ============================================================

function ouvrirExplorateur() {
  const html = HtmlService.createHtmlOutputFromFile('Explorateur')
    .setTitle('🧭 Explorateur')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

function chercherDonneesExplorateur(query, catFilter) {
  _exigerContexteTableur();
  const q = (query || "").toLowerCase().trim();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOM_FEUILLE);
  const fullData = sheet.getDataRange().getValues();
  const toutesLesSorties = _parcourirEntrees(fullData, () => true);
  const results = [];

  toutesLesSorties.forEach(e => {
    if (q && !e.titre.toLowerCase().includes(q)) return;

    let matchFiltre = true;
    if (catFilter) {
      const parts = catFilter.split('_');
      const fCat = parts[0];
      const fStat = parts[1];

      if (fCat === "Films" && !e.cat.includes("Films")) matchFiltre = false;
      if (fCat === "Series" && !e.cat.includes("Séries")) matchFiltre = false;
      if (fCat === "Animes" && !e.cat.includes("Animés")) matchFiltre = false;
      if (fCat === "Jeux" && !e.cat.includes("Jeux")) matchFiltre = false;

      if (fStat === "Vus" && e.statut !== "✔️") matchFiltre = false;
      if (fStat === "Avoir" && e.statut === "✔️") matchFiltre = false;
    }

    if (!matchFiltre) return;

    const dateAff = e.indetermine ? `${e.periodeActuelle} ${e.annee}` : `${(e.jourActuel !== "?" && e.jourActuel) ? e.jourActuel : "?"}/${e.mois + 1}`;

    results.push({
      titre: e.titre,
      statut: e.statut || "❌",
      categorie: e.cat,
      plateforme: e.plateforme,
      dateAff: dateAff,
      ligne: e.ligneAbs,
      colonne: e.colDebut,
      annee: e.annee
    });
  });

  return results;
}

function selectionnerLigne(ligne, col) {
  _exigerContexteTableur();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOM_FEUILLE);
  if (sheet) {
    ss.setActiveSheet(sheet);
    sheet.getRange(ligne, col).activate();
  }
}

// ============================================================
//   AGENDA, GÉNÉRATEUR AVANCÉ, DÉCALAGE, FORMAT CONDITIONNEL
// ============================================================

const NOM_FEUILLE_AGENDA = "🗓️ AGENDA";
const JOURS_FR_LONG  = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi"];
const JOURS_FR_COURT = ["Dim", "Lun", "Mar", "Mer", "Jeu", "Ven", "Sam"];
const MAX_ENTREES_GENERATEUR = 120;

function _pad2(n) { return String(n).padStart(2, "0"); }
function _fmtJJMMAAAA(d) { return `${_pad2(d.getDate())}/${_pad2(d.getMonth() + 1)}/${d.getFullYear()}`; }
function _cleJour(d) { return `${d.getFullYear()}-${_pad2(d.getMonth() + 1)}-${_pad2(d.getDate())}`; }
function _fmtCourt(d) { return `${d.getDate()} ${MOIS_NOMS[d.getMonth()].toLowerCase()}`; }
function _lundiDe(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() - ((d.getDay() + 6) % 7)); }

function _parseDateSouple(s) {
  s = (s || "").toString().trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  return null;
}

function ouvrirGenerateurEpisodes() {
  const html = HtmlService.createTemplateFromFile("GenerateurEpisodes").evaluate()
    .setWidth(680).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, "⚡ Générateur d'épisodes");
}

function _calculerPlanningEpisodes(p) {
  const titre = (p.titre || "").toString().trim();
  if (!titre) throw new Error("Le titre est obligatoire.");
  const premierEp = Math.max(1, parseInt(p.premierEp, 10) || 1);

  const pauses = new Set();
  (p.pauses || "").toString().split(/[,;\n]+/).forEach(s => {
    const d = _parseDateSouple(s);
    if (d) pauses.add(_cleJour(d));
  });

  const seances = [];
  let dernier, manquants = 0;
  const libre = (p.planningLibre || "").toString().trim();

  if (libre) {
    const lignes = [];
    libre.split(/\n+/).forEach(ligne => {
      const l = ligne.trim();
      if (!l) return;
      const m = l.match(/^(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{4})\s*(?:[:=]\s*(\d+))?$/);
      if (!m) throw new Error(`Ligne non comprise : « ${l} » (format attendu : JJ/MM/AAAA : nombre)`);
      lignes.push({ date: _parseDateSouple(m[1]), n: Math.max(1, parseInt(m[2], 10) || 1) });
    });
    if (lignes.length === 0) throw new Error("Le mode libre est vide.");
    lignes.sort((x, y) => x.date - y.date);
    let ep = premierEp;
    lignes.forEach(l => {
      seances.push({ date: l.date, epDebut: ep, epFin: ep + l.n - 1 });
      ep += l.n;
    });
    dernier = ep - 1;
  } else {
    const nbTotal = parseInt(p.nbTotal, 10);
    if (!nbTotal || nbTotal < 1) throw new Error("Le nombre total d'épisodes est obligatoire.");
    const drop = Math.max(1, parseInt(p.drop, 10) || 1);
    const parJour = Math.max(1, parseInt(p.parJour, 10) || 1);
    const intervalle = Math.max(1, parseInt(p.intervalle, 10) || 1);
    const jours = (p.jours || []).map(n => parseInt(n, 10)).filter(n => n >= 0 && n <= 6);
    const debut = _parseDateSouple(p.dateDebut);
    if (!debut) throw new Error("Date de début invalide.");

    dernier = premierEp + nbTotal - 1;
    let ep = premierEp;
    const ajouter = (date, n) => {
      const fin = Math.min(dernier, ep + n - 1);
      seances.push({ date: new Date(date.getTime()), epDebut: ep, epFin: fin });
      ep = fin + 1;
    };

    ajouter(debut, drop);
    const lundi0 = _lundiDe(debut);
    const curseur = new Date(debut.getFullYear(), debut.getMonth(), debut.getDate() + 1);
    let garde = 0;
    while (ep <= dernier && jours.length > 0 && garde++ < 1500) {
      const diffJours = Math.round((curseur - lundi0) / 86400000);
      const idxSemaine = Math.floor(diffJours / 7);
      if (jours.includes(curseur.getDay()) && idxSemaine % intervalle === 0 && !pauses.has(_cleJour(curseur))) {
        ajouter(curseur, parJour);
      }
      curseur.setDate(curseur.getDate() + 1);
    }
    manquants = Math.max(0, dernier - ep + 1);
  }

  const entrees = [];
  seances.forEach(s => {
    const nums = [];
    for (let n = s.epDebut; n <= s.epFin; n++) nums.push(n);
    const groupes = p.regrouper ? [nums] : nums.map(n => [n]);
    groupes.forEach(g => {
      const lib = g.length > 1 ? `${g[0]}-${g[g.length - 1]}` : `${g[0]}`;
      let t = `${titre} - Ep. ${lib}`;
      if (p.marquerFinal && g.includes(dernier)) t += " (Finale)";
      if (p.tagHebdo) t += " [Hebdo]";
      entrees.push({ titre: t, dateStr: _fmtJJMMAAAA(s.date), jour: JOURS_FR_LONG[s.date.getDay()] });
    });
  });

  return { entrees, manquants };
}

function apercuEpisodesAvance(p) {
  _exigerContexteTableur();
  const r = _calculerPlanningEpisodes(p);
  return {
    lignes: r.entrees.map(e => ({ date: e.dateStr, jour: e.jour, titre: e.titre })),
    manquants: r.manquants
  };
}

function genererEpisodesAvance(p) {
  _exigerContexteTableur();
  const r = _calculerPlanningEpisodes(p);
  if (r.entrees.length > MAX_ENTREES_GENERATEUR) {
    throw new Error(`${r.entrees.length} lignes à créer : maximum ${MAX_ENTREES_GENERATEUR} par lancement. Coche "regrouper" ou découpe en plusieurs fois.`);
  }
  let ok = 0;
  const erreurs = [];
  r.entrees.forEach(e => {
    const msg = insererEntree_(e.titre, e.dateStr, p.cat || "🍿 Animés", p.plateforme || "", "", "❌");
    if (msg.startsWith("✅")) ok++;
    else erreurs.push(`${e.dateStr} : ${msg}`);
  });
  mettreEnValeurProchaineSortie();
  _invaliderCacheJSON();
  SpreadsheetApp.flush();
  return { ok, erreurs, manquants: r.manquants };
}

function _baseTitreSerie(titre) {
  return _normSimple(_extraireTitreSansEpisode(titre).replace(/\s*\[.*?\]\s*/g, " ").trim());
}
function _numEpisodeTitre(titre) {
  const m = (titre || "").match(/ep\.?\s*(\d+)/i);
  return m ? parseInt(m[1], 10) : 0;
}

function decalerEpisodesSerie() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOM_FEUILLE);
  if (!sheet) return;

  const repNom = ui.prompt("⏩ Décaler une série", "Titre de la série, sans « Ep. » (ex : Percy Jackson) :", ui.ButtonSet.OK_CANCEL);
  if (repNom.getSelectedButton() !== ui.Button.OK) return;
  const base = _baseTitreSerie(repNom.getResponseText());
  if (!base) return;

  const repDecal = ui.prompt("Décalage en jours", "7 = une semaine plus tard\n-1 = un jour plus tôt", ui.ButtonSet.OK_CANCEL);
  if (repDecal.getSelectedButton() !== ui.Button.OK) return;
  const decalage = parseInt(repDecal.getResponseText().trim(), 10);
  if (isNaN(decalage) || decalage === 0) return ui.alert("❌ Décalage invalide.");

  const repDepuis = ui.prompt("À partir de quel épisode ?", "Numéro (laisser vide = tous les épisodes)", ui.ButtonSet.OK_CANCEL);
  if (repDepuis.getSelectedButton() !== ui.Button.OK) return;
  const depuis = parseInt(repDepuis.getResponseText().trim(), 10) || 0;

  const traites = new Set();
  const erreurs = [];
  let nb = 0, garde = 0;

  while (garde++ < 500) {
    const fullData = sheet.getDataRange().getValues();
    const cibles = _parcourirEntrees(fullData, e =>
      !e.indetermine && e.jourActuel && e.jourActuel !== "?" &&
      !traites.has(e.titre) && _baseTitreSerie(e.titre) === base && _numEpisodeTitre(e.titre) >= depuis
    );
    if (cibles.length === 0) break;

    const c = cibles[0];
    traites.add(c.titre);
    const jour = parseInt(c.jourActuel, 10);
    const nouvelle = new Date(c.annee, c.mois, jour + decalage);
    const ancStr = `${_pad2(jour)}/${_pad2(c.mois + 1)}/${c.annee}`;

    supprimerEntreeInterne_(ss, c, false);
    const msg = insererEntree_(c.titre, _fmtJJMMAAAA(nouvelle), c.cat, c.plateforme, c.lien, c.statut);
    if (msg.startsWith("✅")) nb++;
    else {
      insererEntree_(c.titre, ancStr, c.cat, c.plateforme, c.lien, c.statut);
      erreurs.push(`${c.titre} : ${msg}`);
    }
  }

  mettreEnValeurProchaineSortie();
  _invaliderCacheJSON();
  SpreadsheetApp.flush();
  ui.alert(`⏩ ${nb} épisode(s) décalé(s) de ${decalage} jour(s).` + (erreurs.length ? `\n\n⚠️ Non déplacés :\n${erreurs.join("\n")}` : ""));
}

function genererVueAgenda(silencieux) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(NOM_FEUILLE);
  if (!src) return;
  silencieux = (silencieux === true);

  const fullData = src.getDataRange().getValues();
  const entries = _parcourirEntrees(fullData, () => true);
  const auj = new Date(); auj.setHours(0, 0, 0, 0);
  const DAY = 86400000;

  const datees = [], sansJour = [];
  entries.forEach(e => {
    if (e.indetermine) { sansJour.push({ e, label: e.periodeActuelle || String(e.annee) }); return; }
    const j = parseInt(e.jourActuel, 10);
    if (isNaN(j)) { sansJour.push({ e, label: `${MOIS_NOMS[e.mois]} ${e.annee}` }); return; }
    datees.push({ e, date: new Date(e.annee, e.mois, j) });
  });

  const visibles = datees
    .filter(x => x.date.getTime() >= auj.getTime() - 7 * DAY)
    .sort((a, b) => a.date - b.date || a.e.titre.localeCompare(b.e.titre));

  const rows = [];
  let semCle = null;
  const lundiAuj = _lundiDe(auj).getTime();
  visibles.forEach(x => {
    const l = _lundiDe(x.date);
    if (l.getTime() !== semCle) {
      semCle = l.getTime();
      const fin = new Date(l.getFullYear(), l.getMonth(), l.getDate() + 6);
      rows.push({ type: "week", label: `Semaine du ${_fmtCourt(l)} au ${_fmtCourt(fin)}` + (semCle === lundiAuj ? "    •    CETTE SEMAINE" : "") });
    }
    rows.push({ type: "item", e: x.e, date: x.date });
  });
  if (sansJour.length > 0) {
    rows.push({ type: "indet", label: "📆 DATE INDÉTERMINÉE / SAISON" });
    sansJour.slice(0, 60).forEach(x => rows.push({ type: "item", e: x.e, date: null, labelDate: x.label }));
  }

  let ag = ss.getSheetByName(NOM_FEUILLE_AGENDA);
  if (!ag) { ag = ss.insertSheet(NOM_FEUILLE_AGENDA); ag.setTabColor("#F5B82E"); }
  ag.getRange(1, 1, ag.getMaxRows(), ag.getMaxColumns()).breakApart();
  ag.clear();
  ag.clearConditionalFormatRules();
  ag.setHiddenGridlines(true);

  const NB_COLS = 7;
  const total = 4 + Math.max(1, rows.length);
  if (ag.getMaxRows() < total + 5) ag.insertRowsAfter(ag.getMaxRows(), total + 5 - ag.getMaxRows());
  if (ag.getMaxColumns() < NB_COLS) ag.insertColumnsAfter(ag.getMaxColumns(), NB_COLS - ag.getMaxColumns());

  [110, 120, 340, 130, 170, 60, 100].forEach((w, i) => ag.setColumnWidth(i + 1, w));
  ag.getRange(1, 1, total, NB_COLS).setFontFamily("Roboto").setVerticalAlignment("middle");

  const nbVus = entries.filter(e => e.statut === "✔️").length;
  const nbAVenir = datees.filter(x => x.date >= auj).length;
  const parCat = CATEGORIES.map(c => {
    const n = entries.filter(e => e.cat === c).length;
    return n ? `${c.split(" ")[0]} ${n}` : null;
  }).filter(Boolean).join("    ");
  const maj = Utilities.formatDate(new Date(), FUSEAU_FR, "dd/MM HH:mm");

  ag.setRowHeight(1, 46);
  ag.getRange(1, 1, 1, NB_COLS).merge().setValue("🗓️  AGENDA DES SORTIES")
    .setBackground("#182433").setFontColor("#FFFFFF").setFontSize(20).setFontWeight("bold")
    .setHorizontalAlignment("center");
  ag.setRowHeight(2, 28);
  ag.getRange(2, 1, 1, NB_COLS).merge()
    .setValue(`${nbAVenir} à venir    •    ✔️ ${nbVus}/${entries.length} vus    •    ${parCat}    •    maj ${maj}`)
    .setBackground("#F1F3F5").setFontColor("#68727D").setFontSize(10).setFontStyle("italic")
    .setHorizontalAlignment("center");
  ag.setRowHeight(3, 8);
  ag.setRowHeight(4, 26);
  ag.getRange(4, 1, 1, NB_COLS).setValues([["DATE", "QUAND", "TITRE", "CATÉGORIE", "OÙ VOIR", "VU", "LIEN"]])
    .setBackground("#26354A").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(9)
    .setHorizontalAlignment("center");
  ag.setFrozenRows(4);

  if (rows.length === 0) {
    ag.getRange(5, 1, 1, NB_COLS).merge().setValue("Aucune sortie à venir 🎉")
      .setHorizontalAlignment("center").setFontColor("#68727D");
    if (!silencieux) { ss.setActiveSheet(ag); }
    return;
  }

  const vals = [], bgs = [], fcs = [], fws = [], lines = [], rtvs = [];
  const vide = () => Array(NB_COLS).fill("");
  rows.forEach(r => {
    if (r.type === "week" || r.type === "indet") {
      const row = vide(); row[0] = r.label;
      const bg = r.type === "week" ? "#26354A" : COULEUR_INDETERMINE;
      vals.push(row);
      bgs.push(Array(NB_COLS).fill(bg));
      fcs.push(Array(NB_COLS).fill("#FFFFFF"));
      fws.push(Array(NB_COLS).fill("bold"));
      lines.push(["none"]);
      rtvs.push([SpreadsheetApp.newRichTextValue().setText("").build()]);
      return;
    }

    const e = r.e;
    const passe = r.date ? r.date.getTime() < auj.getTime() : false;
    const vu = e.statut === "✔️";
    const dim = passe || vu;

    let dateLabel, quand = "—";
    if (r.date) {
      dateLabel = `${JOURS_FR_COURT[r.date.getDay()]} ${_pad2(r.date.getDate())}/${_pad2(r.date.getMonth() + 1)}`;
      const diff = Math.round((r.date - auj) / DAY);
      quand = diff === 0 ? "🔥 Aujourd'hui" : diff === 1 ? "Demain" : diff > 1 ? `J-${diff}` : `Sorti (J+${-diff})`;
    } else {
      dateLabel = r.labelDate;
    }

    vals.push([dateLabel, quand, _neutraliserFormule(e.titre), e.cat, _neutraliserFormule(e.plateforme), e.statut || "❌", ""]);

    const bgRow = dim ? "#F3F4F6" : _bgCouleurCat(e.cat);
    const bgDate = dim ? "#C9CED4" : (r.date ? _couleurMois(r.date.getMonth()) : COULEUR_INDETERMINE);
    let bgQuand = bgRow, fcQuand = dim ? "#9AA3AD" : "#20252B", fwQuand = "normal";
    if (r.date && !passe) {
      const diff = Math.round((r.date - auj) / DAY);
      if (diff === 0) { bgQuand = "#D1E7DD"; fcQuand = "#0F5132"; fwQuand = "bold"; }
      else if (diff <= 7) { bgQuand = "#FFF3CD"; fcQuand = "#765600"; fwQuand = "bold"; }
    }
    bgs.push([bgDate, bgQuand, bgRow, bgRow, bgRow, bgRow, bgRow]);
    const fcTxt = dim ? "#9AA3AD" : "#20252B";
    fcs.push(["#FFFFFF", fcQuand, dim ? "#9AA3AD" : "#1A1A2E", fcTxt, fcTxt, fcTxt, fcTxt]);
    fws.push(["bold", fwQuand, "bold", "normal", "normal", "normal", "normal"]);
    lines.push([vu ? "line-through" : "none"]);

    const rtv = SpreadsheetApp.newRichTextValue();
    const lienSur = _urlSure(e.lien);
    if (lienSur) rtv.setText("🔗 Ouvrir").setLinkUrl(lienSur);
    else rtv.setText("");
    rtvs.push([rtv.build()]);
  });

  const n = rows.length;
  const zone = ag.getRange(5, 1, n, NB_COLS);
  zone.setValues(vals).setBackgrounds(bgs).setFontColors(fcs).setFontWeights(fws).setFontSize(11);
  ag.getRange(5, 3, n, 1).setFontLines(lines);
  ag.getRange(5, 7, n, 1).setRichTextValues(rtvs).setFontSize(10);
  ag.getRange(5, 1, n, 2).setHorizontalAlignment("center");
  ag.getRange(5, 4, n, 2).setHorizontalAlignment("center").setFontSize(10);
  ag.getRange(5, 6, n, 2).setHorizontalAlignment("center");
  ag.setRowHeights(5, n, 30);
  zone.setBorder(false, false, false, false, false, true, "#E3E6EA", SpreadsheetApp.BorderStyle.SOLID);

  if (!silencieux) ss.setActiveSheet(ag);
}

function appliquerMiseEnFormeConditionnelle(silencieux) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOM_FEUILLE);
  if (!sheet) return;
  const nbLignes = sheet.getMaxRows() - GRID_START_ROW + 1;
  const colStatut1 = 1 + OFF_STATUT;
  const colStatut2 = 1 + BLOC_WIDTH + OFF_STATUT;
  const lettre = n => String.fromCharCode(64 + n);

  const gardees = sheet.getConditionalFormatRules().filter(r => {
    const c = r.getBooleanCondition();
    return !(c && c.getCriteriaValues && String(c.getCriteriaValues()[0]).indexOf('="✔️"') !== -1);
  });

  const r1 = sheet.getRange(GRID_START_ROW, 2, nbLignes, COLS_PAR_MOIS - 1);
  const r2 = sheet.getRange(GRID_START_ROW, 2 + BLOC_WIDTH, nbLignes, COLS_PAR_MOIS - 1);
  const regle = (range, col) => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$${lettre(col)}${GRID_START_ROW}="✔️"`)
    .setFontColor("#9AA3AD").setStrikethrough(true).setRanges([range]).build();

  sheet.setConditionalFormatRules([...gardees, regle(r1, colStatut1), regle(r2, colStatut2)]);
  if (silencieux !== true) SpreadsheetApp.getUi().alert("✅ Les sorties marquées ✔️ seront désormais grisées et barrées dans la grille.");
}

function majQuotidienneComplete() {
  mettreEnValeurProchaineSortie();
  genererVueAgenda(true);
}

function activerMajQuotidienneComplete() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const h = t.getHandlerFunction();
    if (h === "_majQuotidienne" || h === "majQuotidienneComplete") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("majQuotidienneComplete").timeBased().everyDays(1).atHour(6).create();
  SpreadsheetApp.getUi().alert("✅ Chaque matin à 6h : surbrillance de la prochaine sortie + rafraîchissement de l'agenda.");
}

function exporterCalendrierVersCSV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOM_FEUILLE);
  if (!sheet) return _alerte("❌ Feuille non trouvée");

  const fullData = sheet.getDataRange().getValues();
  const entries = _parcourirEntrees(fullData, () => true);
  const esc = v => `"${(v || "").toString().replace(/"/g, '""')}"`;

  let csvContent = "Titre,Date,Categorie,Plateforme,Lien,Statut\n";

  entries.forEach(e => {
    let dateStr = "";
    if (e.indetermine) {
      dateStr = `${e.periodeActuelle || ""} ${e.annee}`.trim();
    } else {
      const jStr = (e.jourActuel && e.jourActuel !== "?") ? String(e.jourActuel).padStart(2, "0") : "";
      const mStr = String((e.mois || 0) + 1).padStart(2, "0");
      dateStr = jStr ? `${e.annee}-${mStr}-${jStr}` : `${e.annee}-${mStr}-01`;
    }
    csvContent += `${esc(e.titre)},${esc(dateStr)},${esc(e.cat)},${esc(e.plateforme)},${esc(e.lien)},${esc(e.statut || "❌")}\n`;
  });

  DriveApp.createFile("Export_Calendrier_Notion.csv", "\uFEFF" + csvContent, MimeType.CSV);
  _alerte("✅ Fichier Export_Calendrier_Notion.csv créé à la racine de ton Google Drive !");
}