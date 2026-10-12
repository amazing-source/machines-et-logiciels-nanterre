// Seance 6 : mesures reelles faites dans Chrome pilote par Playwright
// (profil temporaire, sans aucun compte personnel).
// Usage : node seance6_mesures.js <tache> <dossier_de_sortie>
//   taches : cookies | clavier | carbone | dnt | atalan | ameli | page <url>
const path = require('path');
const fs = require('fs');
// playwright-core est cherche dans PW_HOME s'il est defini, sinon a cote de ce script
const { chromium } = require(path.resolve(process.env.PW_HOME || __dirname, 'node_modules', 'playwright-core'));

const task = process.argv[2];
const out = process.argv[3] || '.';
fs.mkdirSync(out, { recursive: true });
const save = (name, obj) => fs.writeFileSync(path.join(out, name), typeof obj === 'string' ? obj : JSON.stringify(obj, null, 2), 'utf8');
const sleep = ms => new Promise(r => setTimeout(r, ms));

async function open(opts = {}) {
  const browser = await chromium.launch({ channel: 'chrome', headless: opts.headed ? false : true, args: ['--window-position=40,40', '--mute-audio'] });
  const context = await browser.newContext({ locale: 'fr-FR', viewport: { width: 1366, height: 800 } });
  const page = await context.newPage();
  return { browser, context, page };
}

const site = host => host.replace(/^\./, '').split('.').slice(-2).join('.');

async function cookies() {
  const url = 'https://www.lemonde.fr/';
  const { browser, context, page } = await open();
  const hosts = new Set();
  page.on('request', r => { try { hosts.add(new URL(r.url()).hostname); } catch (e) {} });
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await sleep(8000);
  const snap = async label => {
    const all = await context.cookies();
    const first = all.filter(c => site(c.domain) === 'lemonde.fr');
    const third = all.filter(c => site(c.domain) !== 'lemonde.fr');
    return {
      etat: label, total: all.length, premiere_partie: first.length, tiers: third.length,
      domaines_tiers: [...new Set(third.map(c => site(c.domain)))].sort(),
      noms_premiere_partie: first.map(c => c.name).sort(),
    };
  };
  const avant = await snap('avant tout choix');
  avant.hotes_contactes = hosts.size;
  avant.sites_tiers_contactes = [...new Set([...hosts].map(site))].filter(s => s !== 'lemonde.fr').sort();
  await page.screenshot({ path: path.join(out, 'cookies_avant.png') });
  // bandeau de consentement
  let clique = null;
  for (const frame of page.frames()) {
    for (const txt of ['Accepter et continuer', 'Tout accepter', 'Accepter']) {
      const b = frame.getByRole('button', { name: txt }).first();
      if (await b.count().catch(() => 0)) {
        try { await b.click({ timeout: 4000 }); clique = txt; break; } catch (e) {}
      }
    }
    if (clique) break;
  }
  await sleep(12000);
  await page.mouse.wheel(0, 2500).catch(() => {});
  await sleep(6000);
  const apres = await snap('apres avoir clique sur « ' + clique + ' »');
  apres.hotes_contactes = hosts.size;
  apres.sites_tiers_contactes = [...new Set([...hosts].map(site))].filter(s => s !== 'lemonde.fr').sort();
  await page.screenshot({ path: path.join(out, 'cookies_apres.png') });
  save('cookies.json', { url, date: new Date().toISOString(), bouton: clique, avant, apres });
  console.log('cookies avant:', avant.total, '(tiers', avant.tiers + ') | apres:', apres.total, '(tiers', apres.tiers + ', ', apres.domaines_tiers.length, 'domaines)', '| bouton:', clique);
  await browser.close();
}

async function clavier() {
  const url = 'https://www.parisnanterre.fr/';
  const { browser, page } = await open();
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await sleep(6000);
  const info = () => page.evaluate(() => {
    const e = document.activeElement;
    if (!e || e === document.body) return { tag: 'BODY' };
    const r = e.getBoundingClientRect();
    const cs = getComputedStyle(e);
    const anc = [];
    for (let p = e; p && p !== document.body; p = p.parentElement) {
      const idc = ((p.id || '') + ' ' + (p.className && p.className.toString ? p.className.toString() : '')).toLowerCase();
      if (/cookie|consent|tarteaucitron|didomi|axeptio|rgpd/.test(idc)) { anc.push('bandeau-cookies'); break; }
      if (p.tagName === 'NAV') anc.push('nav');
      if (p.tagName === 'HEADER') anc.push('header');
      if (p.tagName === 'FOOTER') anc.push('footer');
      if (p.tagName === 'MAIN') anc.push('main');
    }
    return {
      tag: e.tagName, texte: (e.innerText || e.getAttribute('aria-label') || e.getAttribute('title') || e.value || '').trim().replace(/\s+/g, ' ').slice(0, 70),
      href: e.getAttribute('href') || '', zone: [...new Set(anc)].join('>'),
      visible: r.width > 0 && r.height > 0 && r.bottom > 0 && r.top < innerHeight && r.right > 0 && r.left < innerWidth,
      contour: cs.outlineStyle !== 'none' && parseFloat(cs.outlineWidth) > 0 ? cs.outlineWidth + ' ' + cs.outlineStyle : '',
      ombre: cs.boxShadow !== 'none' ? 'oui' : '',
    };
  });
  const etapes = [];
  let trouve = null;
  for (let i = 1; i <= 400; i++) {
    await page.keyboard.press('Tab');
    await sleep(60);
    const a = await info();
    a.n = i;
    etapes.push(a);
    if (/contact/i.test(a.texte + ' ' + a.href) && a.zone !== 'bandeau-cookies') { trouve = a; break; }
  }
  await page.screenshot({ path: path.join(out, 'clavier_focus_contact.png') });
  let arrivee = null;
  if (trouve) {
    await page.keyboard.press('Enter');
    await sleep(5000);
    arrivee = { url: page.url(), titre: await page.title() };
    await page.screenshot({ path: path.join(out, 'clavier_page_contact.png') });
  }
  const sansRepere = etapes.filter(e => e.tag !== 'BODY' && !e.contour && !e.ombre).length;
  const invisibles = etapes.filter(e => e.tag !== 'BODY' && !e.visible).length;
  const bandeau = etapes.filter(e => e.zone === 'bandeau-cookies').length;
  const res = {
    url, date: new Date().toISOString(), tabulations: etapes.length, lien_contact: trouve, arrivee,
    premier_arret: etapes[0], arrets_dans_bandeau_cookies: bandeau,
    arrets_sans_contour_ni_ombre: sansRepere, arrets_hors_ecran_ou_invisibles: invisibles,
    lien_evitement: etapes.slice(0, 6).filter(e => /contenu|content|menu|navigation|recherche/i.test(e.texte) && /^#/.test(e.href)),
    etapes,
  };
  save('clavier.json', res);
  console.log('tabulations:', etapes.length, '| contact:', trouve ? trouve.texte + ' -> ' + trouve.href : 'NON TROUVE', '| arrivee:', arrivee && arrivee.url);
  console.log('premier arret:', JSON.stringify(etapes[0]));
  console.log('bandeau:', bandeau, '| sans repere visuel:', sansRepere, '| invisibles:', invisibles);
  await browser.close();
}

async function pageTexte(url, nom, wait = 6000) {
  const { browser, page } = await open();
  const reqs = [];
  page.on('response', r => { if (/json|api/i.test(r.url())) reqs.push(r.url()); });
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }).catch(e => console.log('goto:', e.message));
  await sleep(wait);
  const txt = await page.evaluate(() => document.body ? document.body.innerText : '').catch(() => '');
  const liens = await page.evaluate(() => [...document.querySelectorAll('a')].map(a => (a.innerText || '').trim().replace(/\s+/g, ' ').slice(0, 60) + ' -> ' + a.href).filter(s => s.length > 6)).catch(() => []);
  save(nom + '.txt', 'URL finale : ' + page.url() + '\nTitre : ' + (await page.title().catch(() => '')) + '\n\n' + txt + '\n\n=== LIENS ===\n' + liens.join('\n'));
  await page.screenshot({ path: path.join(out, nom + '.png') }).catch(() => {});
  console.log(nom, '->', page.url(), '|', txt.length, 'caracteres');
  await browser.close();
}

// Do Not Track : documentaire interactif qui se personnalise avec les traces
// du visiteur. On l'ouvre dans une vraie fenetre et on releve a la fois ce
// qu'il affiche et les services qu'il interroge pour nous profiler.
async function dnt() {
  const { browser, page } = await open({ headed: true });
  const appels = [];
  const consoleMsgs = [];
  page.on('console', m => consoleMsgs.push(m.type() + ': ' + m.text().slice(0, 200)));
  page.on('response', async r => {
    const u = r.url();
    if (/\.(js|css|png|jpg|jpeg|gif|svg|woff2?|ttf|mp4|webm|mp3|ogg|ico)(\?|$)/i.test(u)) return;
    let corps = '';
    try { if (/json|text|javascript/i.test(r.headers()['content-type'] || '')) corps = (await r.text()).slice(0, 600); } catch (e) {}
    appels.push({ url: u, statut: r.status(), corps });
  });
  await page.goto('https://donottrack-doc.com/fr/', { waitUntil: 'domcontentloaded', timeout: 60000 }).catch(e => console.log('goto:', e.message));
  await sleep(15000);
  const texte1 = await page.evaluate(() => document.body.innerText).catch(() => '');
  await page.screenshot({ path: path.join(out, 'dnt_1.png') }).catch(() => {});
  await page.goto('https://donottrack-doc.com/fr/episode/1', { waitUntil: 'domcontentloaded', timeout: 60000 }).catch(e => console.log('goto:', e.message));
  await sleep(20000);
  const texte2 = await page.evaluate(() => document.body.innerText).catch(() => '');
  await page.screenshot({ path: path.join(out, 'dnt_2.png') }).catch(() => {});
  // lancement de l'episode 1 (son coupe), releve toutes les 10 secondes
  const releves = [];
  try {
    // les libelles sont en capitales par CSS : on cherche sans tenir compte de la casse
    await page.getByText(/^\s*fermer\s*$/i).first().click({ timeout: 3000 }).catch(() => {});
    const debut = page.locator('a:visible, button:visible').filter({ hasText: /commencer/i }).first();
    await debut.click({ timeout: 8000 });
    for (let k = 1; k <= 9; k++) {
      await sleep(10000);
      const t = await page.evaluate(() => document.body.innerText).catch(() => '');
      releves.push({ seconde: k * 10, texte: t.replace(/\s+/g, ' ').slice(0, 700) });
      await page.screenshot({ path: path.join(out, 'dnt_ep_' + String(k).padStart(2, '0') + '.png') }).catch(() => {});
    }
  } catch (e) { releves.push({ erreur: e.message }); }
  save('dnt.json', { date: new Date().toISOString(), texte_accueil: texte1, texte_episode: texte2, releves, appels, console: consoleMsgs.slice(0, 40) });
  console.log('dnt: accueil', texte1.length, 'car. | episode', texte2.length, 'car. |', appels.length, 'appels |', consoleMsgs.length, 'messages console');
  await browser.close();
}

(async () => {
  if (task === 'cookies') await cookies();
  else if (task === 'clavier') await clavier();
  else if (task === 'dnt') await dnt();
  else if (task === 'page') await pageTexte(process.argv[4], process.argv[5] || 'page', Number(process.argv[6] || 6000));
  else console.log('tache inconnue');
})().catch(e => { console.error('ERREUR', e.message); process.exit(1); });

