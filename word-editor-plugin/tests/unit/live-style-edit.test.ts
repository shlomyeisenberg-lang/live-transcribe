import { describe, expect, it, vi } from 'vitest';
import type { SuperDoc } from 'superdoc';
import { applyLiveStyleEdit, styleTemplateBytes } from '../../src/engine/live-style-edit';
import { changedStyleDefinitions, editStyles, resolveFormat } from '../../src/engine/docx-style-edit';
import { readZip, readEntryText } from '../../src/engine/docx-parts';

const NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML = `<?xml version="1.0"?><w:styles xmlns:w="${NS}"><w:docDefaults><w:pPrDefault><w:pPr><w:bidi/></w:pPr></w:pPrDefault></w:docDefaults><w:latentStyles w:count="12"/><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:jc w:val="left"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Unused"><w:name w:val="Unused"/></w:style><w:style w:type="paragraph" w:styleId="LTR"><w:name w:val="LTR"/><w:basedOn w:val="Normal"/><w:pPr><w:bidi w:val="0"/></w:pPr></w:style></w:styles>`;

describe('live style edits', () => {
  it('imports only changed definitions, with no body snapshot or defaults', async () => {
    const apply = vi.fn(async () => ({ success: true }));
    const host = { activeEditor: { doc: { templates: { apply } } } } as unknown as SuperDoc;
    const result = await applyLiveStyleEdit(host, XML, { kind: 'modify', id: 'Normal', format: { fontSizePt: 18 } }, '12');
    expect(result.outcome.ok).toBe(true);
    const [input, options] = apply.mock.calls[0] as unknown as [{source:{data:string}}, {expectedRevision:string}];
    expect(options).toEqual({ expectedRevision: '12' });
    const bytes = Uint8Array.from(atob(input.source.data), c => c.charCodeAt(0));
    const parts = readZip(bytes)!;
    expect(parts.map(p => p.name)).toEqual(['[Content_Types].xml', '_rels/.rels', 'word/document.xml', 'word/_rels/document.xml.rels', 'word/styles.xml']);
    const styles = await readEntryText(parts.find(p => p.name === 'word/styles.xml')!);
    expect(styles).toContain('w:styleId="Normal"');
    expect(styles).not.toMatch(/docDefaults|latentStyles|Unused|LTR/);
    expect(await readEntryText(parts.find(p => p.name === 'word/document.xml')!)).toContain('<w:body><w:p/></w:body>');
  });

  it('preserves alternate prefixes after the XML prolog', () => {
    const xml = XML.replace(/w:/g, 'a:').replace('xmlns:w=', 'xmlns:a=');
    const edited = editStyles(xml, { kind: 'modify', id: 'Normal', format: { fontSizePt: 18 } });
    if (!edited.ok) throw new Error(edited.message);
    const definitions = changedStyleDefinitions(xml, edited.xml)!;
    expect(definitions).toMatch(/^<a:styles xmlns:a=/);
    expect(definitions).not.toContain('<w:');
    expect(styleTemplateBytes(definitions).length).toBeGreaterThan(0);
  });

  it('resolves bidi inheritance and explicit LTR overrides when reading and writing alignment', () => {
    expect(resolveFormat(XML, 'Normal').alignment).toBe('right');
    expect(resolveFormat(XML, 'LTR').alignment).toBe('left');
    for (const [id, jc] of [['Normal', 'left'], ['LTR', 'right']] as const) {
      const result = editStyles(XML, { kind: 'modify', id, format: { alignment: 'right' } });
      if (!result.ok) throw new Error(result.message);
      const style = result.xml.match(new RegExp(`<w:style[^>]*w:styleId="${id}"[\\s\\S]*?</w:style>`))![0];
      expect(style).toContain(`<w:jc w:val="${jc}"/>`);
      expect(resolveFormat(result.xml, id).alignment).toBe('right');
    }
  });

  it('refuses deletion without asking the engine to replace document content', async () => {
    const apply = vi.fn();
    const host = { activeEditor: { doc: { templates: { apply } } } } as unknown as SuperDoc;
    expect((await applyLiveStyleEdit(host, XML, { kind: 'delete', id: 'Unused' }, '12')).outcome.ok).toBe(false);
    expect(apply).not.toHaveBeenCalled();
  });
});

describe('OOXML named-style toggles (§17.7.3)', () => {
  const base = `<w:styles xmlns:w="${NS}"><w:style w:type="paragraph" w:styleId="Base"><w:name w:val="Base"/><w:rPr><w:b/><w:bCs/><w:i/><w:iCs/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Derived"><w:name w:val="Derived"/><w:basedOn w:val="Base"/><w:rPr><w:b/><w:bCs/><w:i/><w:iCs/></w:rPr></w:style></w:styles>`;
  it('true toggles the inherited value; false is a no-op', () => {
    expect(resolveFormat(base, 'Base')).toMatchObject({ bold:true,italic:true });
    expect(resolveFormat(base, 'Derived')).toMatchObject({ bold:false,italic:false });
    expect(resolveFormat(base.replace(/<w:(b|bCs|i|iCs)\/>/g, '<w:$1 w:val="0"/>'), 'Derived')).toMatchObject({ bold:false,italic:false });
  });
  it('disabling inherited bold/italic writes an on-toggle, and enabling them removes the inversion', () => {
    const inherited = base.replace('<w:rPr><w:b/><w:bCs/><w:i/><w:iCs/></w:rPr></w:style></w:styles>', '</w:style></w:styles>');
    const disabled = editStyles(inherited, {kind:'modify',id:'Derived',format:{bold:false,italic:false}});
    if(!disabled.ok) throw new Error(disabled.message);
    expect(resolveFormat(disabled.xml,'Derived')).toMatchObject({bold:false,italic:false});
    const enabled=editStyles(disabled.xml,{kind:'modify',id:'Derived',format:{bold:true,italic:true}});
    if(!enabled.ok) throw new Error(enabled.message);
    expect(resolveFormat(enabled.xml,'Derived')).toMatchObject({bold:true,italic:true});
    expect(enabled.xml.match(/<w:style[^>]*w:styleId="Derived"[\s\S]*?<\/w:style>/)![0]).not.toMatch(/<w:(b|bCs)(?:\s|\/|>)/);
  });
  it('writes Latin and complex-script toggles against their own inherited values', () => {
    const xml=base.replace('<w:b/><w:bCs/><w:i/><w:iCs/>','<w:b/><w:bCs w:val="0"/>');
    const edited=editStyles(xml,{kind:'modify',id:'Derived',format:{bold:true}});
    if(!edited.ok) throw new Error(edited.message);
    const derived=edited.xml.match(/<w:style[^>]*w:styleId="Derived"[\s\S]*?<\/w:style>/)![0];
    expect(derived).not.toContain('<w:b w:val="0"/>');
    expect(derived).toContain('<w:bCs/>');
    expect(resolveFormat(edited.xml,'Derived').bold).toBe(true);
  });
  it('never treats absent complex-script toggles as Latin toggles', () => {
    const xml=base.replace('<w:b/><w:bCs/><w:i/><w:iCs/>','<w:b/><w:i/>');
    const edited=editStyles(xml,{kind:'modify',id:'Derived',format:{bold:false,italic:false}});
    if(!edited.ok)throw new Error(edited.message);
    const derived=edited.xml.match(/<w:style[^>]*w:styleId="Derived"[\s\S]*?<\/w:style>/)![0];
    expect(derived).toContain('<w:b/><w:bCs w:val="0"/>');
    expect(derived).toContain('<w:i/><w:iCs w:val="0"/>');
    expect(resolveFormat(edited.xml,'Derived')).toMatchObject({bold:false,italic:false});
  });

});


describe('fail-closed browser toggle inheritance', () => {
  const graph = (baseBold: boolean) => `<w:styles xmlns:w="${NS}"><w:style w:type="paragraph" w:styleId="Base"><w:name w:val="Base"/><w:rPr><w:b w:val="${baseBold ? 1 : 0}"/><w:bCs w:val="${baseBold ? 1 : 0}"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Derived"><w:name w:val="Derived"/><w:basedOn w:val="Base"/><w:rPr><w:b/><w:bCs/></w:rPr></w:style></w:styles>`;
  function host() {
    const apply = vi.fn(async () => ({ success:true }));
    return { apply, value:{activeEditor:{doc:{templates:{apply}}}} as unknown as SuperDoc };
  }
  it('refuses a new style that would toggle inherited bold off before any engine mutation', async()=>{
    const h=host();
    const result=await applyLiveStyleEdit(h.value,graph(true),{kind:'create',name:'Off',type:'paragraph',basedOn:'Base',next:null,format:{bold:false}},'1');
    expect(result.outcome).toMatchObject({ok:false,reason:'unsupported-style-inheritance'});expect(h.apply).not.toHaveBeenCalled();
  });
  it('refuses a base edit introducing a Word/renderer disagreement in an untouched descendant',async()=>{
    const h=host();
    const result=await applyLiveStyleEdit(h.value,graph(false),{kind:'modify',id:'Base',format:{bold:true}},'1');
    expect(result.outcome).toMatchObject({ok:false,reason:'unsupported-style-inheritance'});expect(h.apply).not.toHaveBeenCalled();
  });
  it('enabling already inherited bold removes an inversion and remains supported',async()=>{
    const h=host();const result=await applyLiveStyleEdit(h.value,graph(true),{kind:'modify',id:'Derived',format:{bold:true}},'1');
    expect(result.outcome.ok).toBe(true);expect(h.apply).toHaveBeenCalledOnce();
  });
  it('preserves unrelated pre-existing renderer discrepancies when only font size changes',async()=>{
    const h=host();const result=await applyLiveStyleEdit(h.value,graph(true),{kind:'modify',id:'Base',format:{fontSizePt:17}},'1');
    expect(result.outcome.ok).toBe(true);expect(h.apply).toHaveBeenCalledOnce();
  });
});
