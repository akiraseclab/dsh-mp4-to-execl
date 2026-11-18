"""DSH video-to-script utilities. Python 3.10+; ffmpeg/ffprobe on PATH."""
import copy, argparse, hashlib, io, json, os, posixpath, re, subprocess, sys, tempfile, math
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as ET
from PIL import Image, ImageOps
S='http://schemas.openxmlformats.org/spreadsheetml/2006/main'; R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'; X='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'; A='http://schemas.openxmlformats.org/drawingml/2006/main'; P='http://schemas.openxmlformats.org/package/2006/relationships'
NS={'s':S,'r':R,'x':X,'a':A,'p':P}
def q(n,t): return '{'+n+'}'+t
def xml(b): return ET.fromstring(b,parser=ET.XMLParser(resolve_entities=False,no_network=True))
def dump(x): return ET.tostring(x,encoding='UTF-8',xml_declaration=True,standalone=True)
def run(*args): return subprocess.run([str(x) for x in args],check=True,capture_output=True,text=True,encoding='utf-8',errors='replace').stdout
def emit(obj): print(json.dumps(obj,ensure_ascii=False,indent=2))
def sha(path):
 h=hashlib.sha256()
 with open(path,'rb') as f:
  for block in iter(lambda:f.read(1048576),b''): h.update(block)
 return h.hexdigest()
def col(i):
 out=''; i+=1
 while i: i,r=divmod(i-1,26); out=chr(65+r)+out
 return out
def relpath(part): return posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels')
def target(part,t): return t.lstrip('/') if t.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(part),t))
def relationships(z,part):
 p=relpath(part); root=xml(z.read(p)) if p in z.namelist() else ET.Element(q(P,'Relationships'),nsmap={None:P}); return p,root,{e.get('Id'):e for e in root}
def sheets(z):
 root=xml(z.read('xl/workbook.xml')); _,_,rels=relationships(z,'xl/workbook.xml'); return {e.get('name'):target('xl/workbook.xml',rels[e.get(q(R,'id'))].get('Target')) for e in root.xpath('//s:sheet',namespaces=NS)}
def strings(z):
 return [''.join(e.xpath('.//s:t/text()',namespaces=NS)) for e in xml(z.read('xl/sharedStrings.xml'))] if 'xl/sharedStrings.xml' in z.namelist() else []
def value(c,ss):
 if c.find(q(S,'f')) is not None: return '='+str(c.findtext(q(S,'f')))
 if c.get('t')=='inlineStr': return ''.join(c.xpath('./s:is//s:t/text()',namespaces=NS))
 v=c.findtext(q(S,'v')); return ss[int(v)] if c.get('t')=='s' and v is not None else v

def pics(z,part,root):
 _,_,rels=relationships(z,part)
 result=[]
 for d in root.xpath('./s:drawing',namespaces=NS):
  dp=target(part,rels[d.get(q(R,'id'))].get('Target')); dr=xml(z.read(dp)); rp,rr,rd=relationships(z,dp)
  for i,anchor in enumerate(dr):
   origin=anchor.find(q(X,'from')); blips=anchor.xpath('.//a:blip[@r:embed]',namespaces=NS)
   if origin is None or len(blips)!=1: continue
   b=blips[0]; rid=b.get(q(R,'embed')); rel=rd[rid]
   if rel.get('TargetMode')=='External': continue
   cell=col(int(origin.findtext(q(X,'col'))))+str(int(origin.findtext(q(X,'row')))+1)
   result.append({'id':f'{dp}#{i}','cell':cell,'part':dp,'relpart':rp,'rid':rid,'media':target(dp,rel.get('Target')),'anchor':anchor,'blip':b})
 return result

def inspect_book(path):
 with ZipFile(path) as z:
  ss=strings(z); out={}
  for name,part in sheets(z).items():
   root=xml(z.read(part)); out[name]={'cells':{c.get('r'):{'value':value(c,ss),'style':c.get('s'),'formula':c.find(q(S,'f')) is not None} for c in root.xpath('//s:sheetData/s:row/s:c',namespaces=NS)},'merges':root.xpath('//s:mergeCell/@ref',namespaces=NS),'images':[{k:v for k,v in p.items() if k not in ('anchor','blip')} for p in pics(z,part,root)]}
  return out

def set_cell(c,v):
 if c.find(q(S,'f')) is not None: raise ValueError('Refusing to replace formula: '+c.get('r'))
 if v is not None and (isinstance(v,bool) or not isinstance(v,(str,int,float)) or (isinstance(v,float) and not math.isfinite(v))): raise ValueError('Invalid cell value')
 for ch in list(c):
  if ch.tag in (q(S,'v'),q(S,'is')): c.remove(ch)
 if v is None or v=='': c.attrib.pop('t',None); return
 if isinstance(v,(int,float)): c.attrib.pop('t',None); ET.SubElement(c,q(S,'v')).text=str(v); return
 c.set('t','inlineStr'); t=ET.SubElement(ET.SubElement(c,q(S,'is')),q(S,'t')); t.set('{http://www.w3.org/XML/1998/namespace}space','preserve'); t.text=str(v)
def convert_image(path,old,ext):
 with Image.open(io.BytesIO(old)) as ref, Image.open(path) as src:
  size=ref.size; src=ImageOps.exif_transpose(src).convert('RGB'); img=ImageOps.contain(src,size); canvas=Image.new('RGB',size,'white'); canvas.paste(img,((size[0]-img.width)//2,(size[1]-img.height)//2)); o=io.BytesIO(); canvas.save(o,'JPEG' if ext in ('.jpeg','.jpg') else 'PNG',**({'quality':90} if ext in ('.jpeg','.jpg') else {})); return o.getvalue()

def fill(template,plan,output):
 if Path(template).resolve()==Path(output).resolve(): raise ValueError('Output must not overwrite original template')
 if Path(output).exists(): raise ValueError('Output already exists; choose a new path')
 changes={}; added={}; report={'sheets':{},'protected_parts_unchanged':True}
 with ZipFile(template) as z:
  if any(n.startswith('_xmlsignatures/') for n in z.namelist()): raise ValueError('Signed workbook not supported')
  mapping=sheets(z)
  for name,spec in plan['sheets'].items():
   part=mapping[name]; root=xml(z.read(part))
   if root.find(q(S,'sheetProtection')) is not None: raise ValueError('Protected worksheet not supported')
   cells={c.get('r'):c for c in root.xpath('//s:sheetData/s:row/s:c',namespaces=NS)}; allowed=set(spec['allowed_cells']); updates=spec.get('cells',{}); protected=set(spec.get('protected_cells',[]))
   if not set(updates)<=allowed or set(updates)&protected: raise ValueError('Plan attempts to change protected/unapproved cells')
   for addr,v in updates.items():
    if addr not in cells: raise ValueError(f'No existing cell {name}!{addr}; do not alter template geometry')
    set_cell(cells[addr],v)
   for h in root.xpath('./s:hyperlinks/s:hyperlink',namespaces=NS):
    if h.get('ref') in updates: h.getparent().remove(h)
   for hs in root.xpath('./s:hyperlinks',namespaces=NS):
    if len(hs)==0: root.remove(hs)
   if updates: changes[part]=dump(root)
   available={p['id']:p for p in pics(z,part,root)}; replacements=spec.get('images',{}); allowed_images=set(spec.get('allowed_images',[]))
   if not set(replacements)<=allowed_images or not set(replacements)<=set(available): raise ValueError('Unknown or unapproved image anchor')
   roots={}
   for ident,path in replacements.items():
    p=available[ident]; dp=p['part']; rp=p['relpart']
    if dp not in roots: roots[dp]=(xml(changes.get(dp,z.read(dp))),xml(changes.get(rp,z.read(rp))))
    dr,rr=roots[dp]; old=z.read(p['media']); ext=Path(p['media']).suffix.lower()
    if ext not in ('.jpeg','.jpg','.png'): raise ValueError('Unsupported existing image format: '+ext)
    digest=hashlib.sha256((ident+str(path)).encode()).hexdigest()[:18]; media='xl/media/dsh_'+digest+ext; newrid='rIdDSH'+digest
    if media in z.namelist() or media in added: raise ValueError('Media identifier collision')
    added[media]=convert_image(path,old,ext); ET.SubElement(rr,q(P,'Relationship'),Id=newrid,Type=R+'/image',Target=posixpath.relpath(media,posixpath.dirname(dp)))
    idx=int(ident.rsplit('#',1)[1]); dr[idx].xpath('.//a:blip[@r:embed]',namespaces=NS)[0].set(q(R,'embed'),newrid)
   for dp,(dr,rr) in roots.items(): changes[dp]=dump(dr); changes[relpath(dp)]=dump(rr)
   report['sheets'][name]={'updated_cells':len(updates),'updated_images':len(replacements)}
  output=Path(output); output.parent.mkdir(parents=True,exist_ok=True)
  with tempfile.NamedTemporaryFile(dir=output.parent,suffix='.xlsx',delete=False) as f: tmp=Path(f.name)
  try:
   with ZipFile(tmp,'w') as out:
    for info in z.infolist(): out.writestr(copy.copy(info),changes.get(info.filename,z.read(info.filename)))
    for n,b in added.items(): out.writestr(n,b,compress_type=8)
   with ZipFile(tmp) as out:
    if out.testzip(): raise ValueError('ZIP CRC error')
    for n in z.namelist():
     if n not in changes and z.read(n)!=out.read(n): raise ValueError('Unexpected changed part: '+n)
    for n in changes:
     if re.fullmatch(r'xl/drawings/[^/]+\.xml',n):
      orig=xml(z.read(n)); new=xml(out.read(n))
      for node in (orig,new):
       for b in node.xpath('.//a:blip[@r:embed]',namespaces=NS): b.set(q(R,'embed'),'__image__')
      if ET.tostring(orig,method='c14n')!=ET.tostring(new,method='c14n'): raise ValueError('Drawing geometry changed: '+n)
   before=inspect_book(template); after=inspect_book(tmp)
   for name,old in before.items():
    new=after[name]; spec=plan['sheets'].get(name,{}); editable=set(spec.get('cells',{}))
    if old['merges']!=new['merges'] or old['cells'].keys()!=new['cells'].keys(): raise ValueError('Cell layout changed')
    if any(old['cells'][a]['style']!=new['cells'][a]['style'] for a in old['cells']): raise ValueError('Style changed')
    if any(old['cells'][a]!=new['cells'][a] for a in old['cells'] if a not in editable): raise ValueError('Unapproved cell changed')
    if [(p['id'],p['cell']) for p in old['images']]!=[(p['id'],p['cell']) for p in new['images']]: raise ValueError('Image anchors changed')
   os.replace(tmp,output)
  finally: tmp.unlink(missing_ok=True)
 report['output']=str(output.resolve()); return report

def prepare(source,outdir,interval):
 d=Path(outdir).resolve(); d.mkdir(parents=True,exist_ok=True); f=Path(source).expanduser()
 if not f.is_file():
  match=re.search(r'https?://[^\s<>"“”]+',source)
  if not match: raise ValueError('Provide a local video or accessible HTTP(S) URL')
  url=match.group(0).rstrip('，。；,;)）'); dl=d/'download'; dl.mkdir(exist_ok=True)
  cmd=[sys.executable,'-m','yt_dlp','--no-playlist','--no-config','--no-progress','--merge-output-format','mp4','-f','bv*+ba/b','-o',str(dl/'source.%(ext)s'),'--print','after_move:filepath','--',url]
  result=run(*cmd).strip().splitlines()
  if not result: raise ValueError('Downloader returned no usable file')
  f=Path(result[-1])
  if not f.is_file(): raise ValueError('Downloader returned no usable file')
 f=f.resolve(); meta=json.loads(run('ffprobe','-v','error','-show_streams','-show_format','-of','json',f)); stream=next(s for s in meta['streams'] if s['codec_type']=='video'); duration=float(meta['format']['duration']); interval=max(float(interval),0.25)
 frames=d/'frames'; frames.mkdir(exist_ok=True); manifest=[]
 for i in range(min(600,max(1,int(duration/interval)+1))):
  t=round(i*interval,3)
  if t>=duration: break
  fp=frames/f'{i:04d}_{t:09.3f}.jpg'; run('ffmpeg','-v','error','-y','-ss',t,'-i',f,'-frames:v','1','-vf','scale=720:-2','-q:v','3',fp)
  if fp.exists(): manifest.append({'time':t,'path':str(fp)})
 hasaudio=any(s['codec_type']=='audio' for s in meta['streams']); audio=None
 if hasaudio: audio=d/'audio.wav'; run('ffmpeg','-v','error','-y','-i',f,'-vn','-ac','1','-ar','16000',audio)
 result={'source':str(f),'sha256':sha(f),'duration':duration,'width':stream['width'],'height':stream['height'],'audio':str(audio) if audio else None,'subtitle_streams':[s['index'] for s in meta['streams'] if s['codec_type']=='subtitle'],'frames':manifest,'sample_interval':interval,'sampling_truncated':duration/interval>600}
 (d/'manifest.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8'); return result

def screenshot(video,seconds,output):
 path=Path(output); path.parent.mkdir(parents=True,exist_ok=True)
 if path.exists(): raise ValueError('Screenshot path already exists')
 run('ffmpeg','-v','error','-y','-ss',float(seconds),'-i',video,'-frames:v','1','-q:v','2',output)
 if not path.is_file(): raise ValueError('No screenshot produced; verify time range')
 return str(path.resolve())
def main():
 p=argparse.ArgumentParser(); s=p.add_subparsers(dest='cmd',required=True)
 t=s.add_parser('prepare'); t.add_argument('source'); t.add_argument('outdir'); t.add_argument('--interval',type=float,default=1)
 t=s.add_parser('inspect'); t.add_argument('template'); t.add_argument('--output')
 t=s.add_parser('frame'); t.add_argument('video'); t.add_argument('seconds',type=float); t.add_argument('output')
 t=s.add_parser('fill'); t.add_argument('template'); t.add_argument('plan'); t.add_argument('output')
 a=p.parse_args()
 try:
  if a.cmd=='prepare': result=prepare(a.source,a.outdir,a.interval)
  elif a.cmd=='inspect':
   result=inspect_book(a.template)
   if a.output: Path(a.output).write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
  elif a.cmd=='frame': result=screenshot(a.video,a.seconds,a.output)
  else: result=fill(a.template,json.loads(Path(a.plan).read_text(encoding='utf-8-sig')),a.output)
  emit(result)
 except Exception as e:
  print('ERROR: '+str(e),file=sys.stderr)
  if isinstance(e,subprocess.CalledProcessError): print(e.stderr,file=sys.stderr)
  sys.exit(1)
if __name__=='__main__': main()
