import JSZip from 'jszip';

const FONTS_LIB = '/SiteAssets/Fonts';
const REL_FONT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font';
const CT_ODTTF = 'application/vnd.openxmlformats-officedocument.obfuscatedFont';

const STYLE = new Set([
  'italic', 'oblique', 'bold', 'light', 'medium', 'regular', 'normal', 'roman',
  'black', 'thin', 'book', 'heavy', 'extralight', 'ultralight', 'semibold', 'demibold', 'extrabold', 'ultrabold'
]);
const ITALIC = new Set(['italic', 'oblique']);
const BOLD = new Set(['bold', 'black', 'heavy', 'extrabold', 'ultrabold', 'semibold', 'demibold']);

type Slot = 'Regular' | 'Bold' | 'Italic' | 'BoldItalic';
const SLOTS: Slot[] = ['Regular', 'Bold', 'Italic', 'BoldItalic'];

export interface IDocxFontPreprocessResult {
  proceed: boolean;
  processedBlob: Blob;
  injectedFonts: string[];
  unresolvedFonts: string[];
}

interface ILibItem { name: string; serverRelativeUrl: string }
interface IEmbed { rid: string; key: string }
interface IEntry { fontName: string; existing: Partial<Record<Slot, IEmbed>> }

interface IPreprocessParams {
  docxBlob: Blob;
  fileName: string;
  sharePointWebUrl: string;
  confirm: (title: string, message: string) => Promise<boolean>;
}

const tokenize = (v: string): string[] => v
  .normalize('NFKD')
  .replace(/[\u0300-\u036f]/g, '')
  .replace(/[\u2010-\u2015]/g, '-')
  .toLowerCase()
  .split(/[^a-z0-9]+/g)
  .filter(Boolean);

const esc = (v: string): string => v.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

const origin = (webUrl: string): string => new URL(webUrl).origin;

const listFonts = async (o: string): Promise<ILibItem[]> => {
  const u = `${o}/_api/web/GetFolderByServerRelativePath(decodedurl='${FONTS_LIB}')/Files?$select=Name,ServerRelativeUrl`;
  const r = await fetch(u, { method: 'GET', headers: { Accept: 'application/json;odata=nometadata' }, credentials: 'include' });
  if (!r.ok) throw new Error(`Font library list failed (${r.status}).`);
  const j = await r.json() as { value?: Array<{ Name: string; ServerRelativeUrl: string }> };
  return (j.value ?? []).filter((f) => /\.(ttf|otf)$/i.test(f.Name)).map((f) => ({ name: f.Name, serverRelativeUrl: f.ServerRelativeUrl }));
};

type Split = { fam: string; sty: Set<string> };
const split = (v: string): Split => {
  const fam: string[] = [];
  const sty = new Set<string>();
  for (const t of tokenize(v)) {
    if (STYLE.has(t)) sty.add(t);
    else fam.push(t);
  }
  return { fam: fam.sort().join('|'), sty };
};

const slotFor = (entry: Split, cand: Split): Slot | undefined => {
  if (entry.fam !== cand.fam) return undefined;
  let it = false, bd = false, other = false;
  cand.sty.forEach((s) => {
    if (entry.sty.has(s)) return;
    if (ITALIC.has(s)) it = true;
    else if (BOLD.has(s)) bd = true;
    else other = true;
  });
  if (other) return undefined;
  if (!it && !bd) return 'Regular';
  if (it && !bd) return 'Italic';
  if (!it && bd) return 'Bold';
  return 'BoldItalic';
};

const matchSlots = (fontName: string, lib: ILibItem[]): Partial<Record<Slot, ILibItem>> => {
  const e = split(fontName);
  if (!e.fam) return {};
  const best: Partial<Record<Slot, { it: ILibItem; ttf: boolean }>> = {};
  for (const it of lib) {
    const c = split(it.name.replace(/\.(ttf|otf)$/i, ''));
    const sl = slotFor(e, c);
    if (!sl) continue;
    const ttf = /\.ttf$/i.test(it.name);
    const cur = best[sl];
    if (!cur || (ttf && !cur.ttf)) best[sl] = { it, ttf };
  }
  const out: Partial<Record<Slot, ILibItem>> = {};
  (Object.keys(best) as Slot[]).forEach((k) => { out[k] = best[k]!.it; });
  return out;
};

const fetchFont = async (item: ILibItem, o: string): Promise<ArrayBuffer | null> => {
  const escPath = item.serverRelativeUrl.replace(/'/g, "''");
  const api = `${o}/_api/web/GetFileByServerRelativePath(decodedurl='${escPath}')/$value`;
  const direct = new URL(item.serverRelativeUrl, o).toString();
  for (const u of [api, direct]) {
    const r = await fetch(u, { method: 'GET', headers: { Accept: 'application/octet-stream' }, credentials: 'include' });
    if (r.ok) return r.arrayBuffer();
  }
  return null;
};

const isSfnt = (b: Uint8Array): boolean =>
  b.length >= 4 && (
    (b[0] === 0 && b[1] === 1 && b[2] === 0 && b[3] === 0) ||
    (b[0] === 0x4f && b[1] === 0x54 && b[2] === 0x54 && b[3] === 0x4f) ||
    (b[0] === 0x74 && b[1] === 0x72 && b[2] === 0x75 && b[3] === 0x65) ||
    (b[0] === 0x74 && b[1] === 0x79 && b[2] === 0x70 && b[3] === 0x31)
  );

const parseEmbed = (inner: string, tag: string): IEmbed | undefined => {
  const a = inner.match(new RegExp(`<${tag}\\b[^>]*\\br:id="([^"]+)"[^>]*\\bw:fontKey="([^"]+)"`, 'i'));
  const b = inner.match(new RegExp(`<${tag}\\b[^>]*\\bw:fontKey="([^"]+)"[^>]*\\br:id="([^"]+)"`, 'i'));
  const m = a || b;
  if (!m) return undefined;
  const rid = m[1].startsWith('rId') ? m[1] : m[2];
  const key = m[1].startsWith('rId') ? m[2] : m[1];
  return { rid, key };
};

const extractEntries = (xml: string): IEntry[] => {
  const out: IEntry[] = [];
  const re = /<w:font\b[^>]*\bw:name="([^"]+)"[^>]*>([\s\S]*?)<\/w:font>/gi;
  let m = re.exec(xml);
  while (m) {
    const existing: Partial<Record<Slot, IEmbed>> = {};
    for (const s of SLOTS) {
      const e = parseEmbed(m[2], `w:embed${s}`);
      if (e) existing[s] = e;
    }
    out.push({ fontName: m[1], existing });
    m = re.exec(xml);
  }
  return out;
};

const newGuid = (): string => {
  const b = new Uint8Array(16);
  crypto.getRandomValues(b);
  b[6] = (b[6] & 0x0f) | 0x40;
  b[8] = (b[8] & 0x3f) | 0x80;
  const h = Array.from(b).map((x) => x.toString(16).padStart(2, '0').toUpperCase()).join('');
  return `{${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20, 32)}}`;
};

const obfuscate = (data: Uint8Array, guid: string): Uint8Array => {
  const hex = guid.replace(/[{}\-]/g, '');
  const raw = new Uint8Array(16);
  for (let i = 0; i < 16; i++) raw[i] = parseInt(hex.slice(i * 2, i * 2 + 2), 16);
  const out = new Uint8Array(data);
  const n = Math.min(32, out.length);
  for (let i = 0; i < n; i++) out[i] ^= raw[15 - (i % 16)];
  return out;
};

const upsertEmbed = (xml: string, name: string, slot: Slot, rid: string, key: string): string => {
  const br = new RegExp(`<w:font\\b[^>]*\\bw:name="${esc(name)}"[^>]*>[\\s\\S]*?<\\/w:font>`, 'i');
  const match = xml.match(br);
  if (!match) return xml;
  const tag = `w:embed${slot}`;
  const node = `<${tag} r:id="${rid}" w:fontKey="${key}"/>`;
  const tr = new RegExp(`<${tag}\\b[^>]*\\/>`, 'gi');
  let block = match[0]
    .replace(/<w:notTrueType\s*\/>/gi, '')
    .replace(/<w:notTrueType>\s*<\/w:notTrueType>/gi, '');
  block = block.match(tr) ? block.replace(tr, node) : block.replace('</w:font>', `${node}</w:font>`);
  return xml.replace(match[0], block);
};

const patchContentTypes = (xml: string): string =>
  xml
    .replace(/<Default\s+Extension="ttf"\s+ContentType="[^"]+"\s*\/>/gi, '')
    .replace(/<Default\s+Extension="otf"\s+ContentType="[^"]+"\s*\/>/gi, '')
    .replace(/<Default\s+Extension="odttf"\s+ContentType="[^"]+"\s*\/>/gi, '')
    .replace('</Types>', `<Default Extension="odttf" ContentType="${CT_ODTTF}"/></Types>`);

const patchRels = (existing: string | null, rels: Array<{ id: string; target: string }>): string => {
  if (!existing?.trim()) {
    const nodes = rels.map((r) => `<Relationship Id="${r.id}" Type="${REL_FONT}" Target="${r.target}"/>`).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${nodes}</Relationships>`;
  }
  let x = existing;
  for (const r of rels) {
    const rx = new RegExp(`<Relationship\\b[^>]*\\bId="${esc(r.id)}"[^>]*\\/?>`, 'i');
    const node = `<Relationship Id="${r.id}" Type="${REL_FONT}" Target="${r.target}"/>`;
    x = rx.test(x) ? x.replace(rx, node) : x.replace('</Relationships>', `${node}</Relationships>`);
  }
  return x;
};

const stripTrash = (zip: JSZip): void => {
  const paths: string[] = [];
  zip.forEach((p) => { if (p.toLowerCase().startsWith('[trash]')) paths.push(p); });
  paths.forEach((p) => zip.remove(p));
  zip.remove('[trash]/');
  zip.remove('[trash]');
};

export const preprocessDocxFonts = async (p: IPreprocessParams): Promise<IDocxFontPreprocessResult> => {
  const { docxBlob, fileName, sharePointWebUrl, confirm } = p;
  const o = origin(sharePointWebUrl);
  const zip = await JSZip.loadAsync(docxBlob);
  const ftXml = (await zip.file('word/fontTable.xml')?.async('string')) ?? '';
  if (!ftXml) return { proceed: true, processedBlob: docxBlob, injectedFonts: [], unresolvedFonts: [] };

  const entries = extractEntries(ftXml);
  if (!entries.length) return { proceed: true, processedBlob: docxBlob, injectedFonts: [], unresolvedFonts: [] };

  let lib: ILibItem[] = [];
  try {
    lib = await listFonts(o);
  } catch {
    const proceed = await confirm(
      'Font Library Unavailable',
      `The shared font library (${FONTS_LIB}) could not be accessed for "${fileName}". Continue without font pre-processing?`
    );
    return { proceed, processedBlob: docxBlob, injectedFonts: [], unresolvedFonts: [] };
  }

  const relsXml = (await zip.file('word/_rels/fontTable.xml.rels')?.async('string')) ?? null;
  const taken = new Set<string>();
  if (relsXml) {
    let rm: RegExpExecArray | null;
    const rre = /Id="(rId\d+)"/g;
    while ((rm = rre.exec(relsXml))) taken.add(rm[1]);
  }
  entries.forEach((e) => {
    (Object.keys(e.existing) as Slot[]).forEach((s) => { const x = e.existing[s]; if (x) taken.add(x.rid); });
  });
  const nextRid = (): string => {
    let i = 1;
    while (taken.has(`rId${i}`)) i++;
    const r = `rId${i}`;
    taken.add(r);
    return r;
  };

  let table = ftXml;
  const rels: Array<{ id: string; target: string }> = [];
  const injected: string[] = [];
  const missing: string[] = [];
  let idx = 0;

  for (const e of entries) {
    const slots = matchSlots(e.fontName, lib);
    const keys = Object.keys(slots) as Slot[];
    if (!keys.length) continue;

    let any = false;
    for (const s of keys) {
      const item = slots[s]!;
      const buf = await fetchFont(item, o);
      if (!buf) {
        missing.push(`${e.fontName} (${s})`);
        continue;
      }
      const raw = new Uint8Array(buf);
      if (!isSfnt(raw)) {
        missing.push(`${e.fontName} (${s})`);
        continue;
      }
      const ex = e.existing[s];
      const rid = ex?.rid ?? nextRid();
      const key = ex?.key ?? newGuid();
      table = upsertEmbed(table, e.fontName, s, rid, key);
      idx++;
      const name = `font${idx}.odttf`;
      zip.file(`word/fonts/${name}`, obfuscate(raw, key));
      rels.push({ id: rid, target: `fonts/${name}` });
      any = true;
    }
    if (any) injected.push(e.fontName);
  }

  if (rels.length) {
    zip.file('word/fontTable.xml', table, { compression: 'DEFLATE' });
    zip.file('word/_rels/fontTable.xml.rels', patchRels(relsXml, rels), { compression: 'STORE' });
    const ct = zip.file('[Content_Types].xml');
    if (ct) zip.file('[Content_Types].xml', patchContentTypes(await ct.async('string')), { compression: 'STORE' });
  }

  stripTrash(zip);

  const processedBlob = rels.length
    ? await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } })
    : docxBlob;

  if (missing.length) {
    const proceed = await confirm(
      'Missing Fonts Detected',
      `The document "${fileName}" references custom fonts expected in ${FONTS_LIB} but could not be loaded: ${missing.join(', ')}. Continue anyway?`
    );
    return { proceed, processedBlob, injectedFonts: injected, unresolvedFonts: missing };
  }

  return { proceed: true, processedBlob, injectedFonts: injected, unresolvedFonts: [] };
};
