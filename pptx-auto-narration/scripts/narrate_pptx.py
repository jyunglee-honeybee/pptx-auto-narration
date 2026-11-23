#!/usr/bin/env python3
"""Turn a PowerPoint deck whose speaker notes hold a click-by-click script into a
self-running deck: TTS narration embedded per slide, every click replaced by a timed
"With Previous" effect, slides advancing automatically.

Subcommands
  inspect  deck.pptx                      click groups / effects / note cues per slide
  parse    deck.pptx -o segs.json         split notes into narration segments
  samples  --text "..." --voices v1,v2    comparison MP3s (Azure)
  tts      segs.json -o segs_audio.json   synthesize every segment (Azure), measure durations
  build    deck.pptx segs_audio.json -o out.pptx [--cues cues.json]

Azure credentials: env AZURE_SPEECH_KEY, AZURE_SPEECH_REGION (or ~/.azkey/key, ~/.azkey/region).
"""
import argparse, hashlib, json, os, re, sys, shutil, subprocess, tempfile, time, html, wave, zipfile, urllib.request
from lxml import etree

P='http://schemas.openxmlformats.org/presentationml/2006/main'
A='http://schemas.openxmlformats.org/drawingml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P14='http://schemas.microsoft.com/office/powerpoint/2010/main'
PR='http://schemas.openxmlformats.org/package/2006/relationships'
CT='http://schemas.openxmlformats.org/package/2006/content-types'
ns={'p':P,'a':A,'r':R}

# ---------------------------------------------------------------- package helpers
def unpack(pptx, d):
    with zipfile.ZipFile(pptx) as z: z.extractall(d)
def pack(d, out):
    if os.path.exists(out): os.remove(out)
    with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as z:
        z.write(os.path.join(d,'[Content_Types].xml'),'[Content_Types].xml')
        for root,_,files in os.walk(d):
            for f in sorted(files):
                full=os.path.join(root,f); rel=os.path.relpath(full,d).replace(os.sep,'/')
                if rel=='[Content_Types].xml': continue
                z.write(full,rel, zipfile.ZIP_STORED if rel.endswith(('.m4a','.jpg','.jpeg','.mp4')) else zipfile.ZIP_DEFLATED)
def slide_order(d):
    pres=etree.parse(d+'/ppt/presentation.xml')
    rmap={r.get('Id'):r.get('Target') for r in etree.parse(d+'/ppt/_rels/presentation.xml.rels').getroot()}
    return [rmap[s.get('{%s}id'%R)] for s in pres.findall('.//p:sldId',ns)]
def rels_path(path): return os.path.join(os.path.dirname(path),'_rels',os.path.basename(path)+'.rels')
def notes_text(d,target):
    sp=d+'/ppt/'+target
    if not os.path.exists(rels_path(sp)): return ''
    for r in etree.parse(rels_path(sp)).getroot():
        if r.get('Type').endswith('/notesSlide'):
            n=etree.parse(os.path.normpath(os.path.join(os.path.dirname(sp),r.get('Target'))))
            for shp in n.findall('.//p:sp',ns):
                ph=shp.find('.//p:ph',ns)
                if ph is not None and ph.get('type')=='body':
                    return '\n'.join(''.join(x.itertext()) for x in shp.findall('.//a:p',ns))
    return ''

# ---------------------------------------------------------------- animation model
def num(v):
    try: return int(v)
    except (TypeError,ValueError): return None
def cdelay(ctn):
    c=ctn.find('p:stCondLst/p:cond',ns)
    return num(c.get('delay')) or 0 if c is not None else 0
def ctn_end(ctn):
    """latest end (ms, relative to ctn start) of anything inside ctn"""
    best=num(ctn.get('dur')) or 0
    for ch in ctn.iter('{%s}cTn'%P):
        if ch is ctn: continue
        s=0; node=ch
        while node is not ctn:
            if node.tag=='{%s}cTn'%P: s+=cdelay(node)
            node=node.getparent()
        best=max(best, s+(num(ch.get('dur')) or 0))
    return best
def groups(root):
    """top-level click groups of the main sequence -> list of dicts"""
    ms=root.find('.//p:cTn[@nodeType="mainSeq"]',ns)
    if ms is None: return None,[]
    out=[]
    for par in ms.find('p:childTnLst',ns).findall('p:par',ns):
        ctn=par.find('p:cTn',ns)
        auto=any(c.get('evt')=='onBegin' for c in ctn.findall('p:stCondLst/p:cond',ns))
        effs=[]
        for sp in ctn.find('p:childTnLst',ns).findall('p:par',ns):
            sctn=sp.find('p:cTn',ns); sd=cdelay(sctn)
            for e in sctn.find('p:childTnLst',ns).findall('p:par',ns):
                ectn=e.find('p:cTn',ns)
                st=sd+cdelay(ectn)
                tg=ectn.find('.//p:spTgt',ns)
                effs.append({'par':e,'ctn':ectn,'start':st,'end':st+ctn_end(ectn),
                             'spid':tg.get('spid') if tg is not None else None,'cls':ectn.get('presetClass')})
        out.append({'par':par,'auto':auto,'effects':effs,'end':max([e['end'] for e in effs],default=0)})
    return ms,out
def shape_labels(root):
    lab={}
    for c in root.iter('{%s}cNvPr'%P):
        sp=c.getparent().getparent(); txt=''.join(sp.itertext()).strip().replace('\n',' ')
        lab[c.get('id')]=(c.get('name') or '')+('|'+txt[:16] if txt else '')
    return lab

# ---------------------------------------------------------------- notes parsing
CLICK_RE=re.compile(r'▶\s*(?:클릭|[Cc]lick)\s*(\d+)(?:\s*\((?:애니메이션|[Aa]nimation)\s*([\d.]+)\s*(?:초|s)\))?')
CUE_RE=re.compile(r'【\s*([\d.]+)\s*(?:초|s)[^】]*】\s*(.*)')
def parse_notes(text):
    """Notes format: lines before the first '▶ 클릭 1' are spoken when the slide appears;
    '▶ 클릭 N' / '▶ Click N' starts click N; '【X초…】' / '【Xs…】 text' = sentence tied to an effect X s after the click;
    skipped: lines starting with ※, ━━ headers (a '사용법' header block is skipped whole), '(보충' marker lines."""
    clicks={0:[]}; cur=0; in_usage=False
    for ln in text.split('\n'):
        s=ln.strip()
        if not s: continue
        if s.startswith('━━'):
            in_usage='사용법' in s; continue
        if in_usage or s.startswith('※') or s.startswith('(보충'): continue
        m=CLICK_RE.match(s)
        if m: cur=int(m.group(1)); clicks[cur]=[]; continue
        m=CUE_RE.match(s)
        if m: clicks[cur].append({'cue':float(m.group(1)),'text':m.group(2)}); continue
        if clicks[cur]: clicks[cur][-1]['text']+=s
        else: clicks[cur].append({'cue':0.0,'text':s})
    return {k:v for k,v in clicks.items() if v or k>0}

def cmd_inspect(a):
    d=tempfile.mkdtemp(); unpack(a.deck,d); bad=0
    for i,t in enumerate(slide_order(d),1):
        root=etree.parse(d+'/ppt/'+t).getroot(); _,gs=groups(root); lab=shape_labels(root)
        nclick=sum(1 for g in gs if not g['auto']); notes=parse_notes(notes_text(d,t))
        nk=max(notes) if notes else 0
        flag='' if nk==nclick else '   <<< 노트의 클릭 수와 애니메이션 클릭 수가 다름'
        bad+=bool(flag)
        print(f'=== 슬라이드 {i}: 클릭 {nclick}, 자동시작 묶음 {sum(g["auto"] for g in gs)}, 노트 클릭 {nk}{flag}')
        k=0
        for g in gs:
            if not g['auto']: k+=1
            print(f'  {"자동" if g["auto"] else "클릭"+str(k)}  (효과 끝 {g["end"]/1000:.1f}s)')
            if a.verbose:
                for e in g['effects']:
                    print(f'      {e["start"]:>6}ms  {e["cls"] or "-":9} spid {e["spid"]:>4}  {lab.get(e["spid"],"")}')
            for sg in notes.get(k if not g['auto'] else -1,[]):
                if sg['cue']>0 and a.verbose: print(f'      cue {sg["cue"]}s  {sg["text"][:30]}')
    shutil.rmtree(d); print('\n불일치 슬라이드 수:',bad)

def cmd_parse(a):
    d=tempfile.mkdtemp(); unpack(a.deck,d); out=[]
    for i,t in enumerate(slide_order(d),1):
        out.append({'slide':i,'file':t,'clicks':{str(k):v for k,v in parse_notes(notes_text(d,t)).items()}})
    shutil.rmtree(d)
    json.dump(out,open(a.o,'w'),ensure_ascii=False,indent=1)
    n=sum(len(v) for s in out for v in s['clicks'].values()); c=sum(len(sg['text']) for s in out for v in s['clicks'].values() for sg in v)
    print(f'{len(out)} slides, {n} segments, {c} characters -> {a.o}')

# ---------------------------------------------------------------- Azure TTS
def azure_cred():
    k=os.environ.get('AZURE_SPEECH_KEY') or open(os.path.expanduser('~/.azkey/key')).read().strip()
    r=os.environ.get('AZURE_SPEECH_REGION') or open(os.path.expanduser('~/.azkey/region')).read().strip()
    return k,r
def azure_tts(text,out,voice,style=None,rate=None,fmt='riff-24khz-16bit-mono-pcm',lang=None):
    key,region=azure_cred()
    lang=lang or '-'.join(voice.split('-')[:2])          # e.g. zh-CN-YunyangNeural -> zh-CN
    body=html.escape(text)
    if rate: body=f'<prosody rate="{rate}">{body}</prosody>'
    if style: body=f'<mstts:express-as style="{style}">{body}</mstts:express-as>'
    ssml=(f'<speak version="1.0" xmlns="http://www.w3.org/2001/10/synthesis" xmlns:mstts="https://www.w3.org/2001/mstts" '
          f'xml:lang="{lang}"><voice name="{voice}">{body}</voice></speak>')
    for attempt in range(6):
        req=urllib.request.Request(f'https://{region}.tts.speech.microsoft.com/cognitiveservices/v1',data=ssml.encode(),method='POST',
            headers={'Ocp-Apim-Subscription-Key':key,'Content-Type':'application/ssml+xml','X-Microsoft-OutputFormat':fmt,'User-Agent':'pptx-narration'})
        try:
            with urllib.request.urlopen(req,timeout=120) as r: open(out,'wb').write(r.read()); return
        except urllib.error.HTTPError as e:
            if e.code in (429,500,502,503): time.sleep(10*(attempt+1)); continue
            raise SystemExit(f'Azure HTTP {e.code}: {e.read()[:300]!r}')
    raise SystemExit('Azure: too many retries')
def cmd_samples(a):
    os.makedirs(a.outdir,exist_ok=True)
    for v in a.voices.split(','):
        v=v.strip(); out=os.path.join(a.outdir,f'{a.prefix}{v.replace(":","_")}.mp3')
        azure_tts(a.text,out,v,a.style,a.rate,fmt='audio-48khz-192kbitrate-mono-mp3',lang=a.lang); print(out); time.sleep(3)
def cmd_tts(a):
    segs=json.load(open(a.segs)); os.makedirs(a.outdir,exist_ok=True); calls=0
    for s in segs:
        for k,v in s['clicks'].items():
            for j,seg in enumerate(v):
                h=hashlib.sha1(f"{a.voice}|{a.style}|{a.rate}|{seg['text']}".encode()).hexdigest()[:8]   # text/voice change -> new file
                out=os.path.join(a.outdir,f"s{s['slide']:02d}_c{int(k):02d}_{j}_{h}.wav"); seg['wav']=out
                if os.path.exists(out) and os.path.getsize(out)>1000 and not a.force: continue
                azure_tts(seg['text'],out,a.voice,a.style,a.rate,lang=a.lang); calls+=1
                print(out,flush=True); time.sleep(3.2)        # free tier: ~20 requests/min
    for s in segs:
        for v in s['clicks'].values():
            for seg in v:
                w=wave.open(seg['wav']); seg['dur']=w.getnframes()/w.getframerate(); w.close()
    json.dump(segs,open(a.o,'w'),ensure_ascii=False,indent=1)
    print(f'{calls} new calls; speech total {sum(seg["dur"] for s in segs for v in s["clicks"].values() for seg in v)/60:.1f} min -> {a.o}')

# ---------------------------------------------------------------- build
def E(tag,**at):
    pre,t=tag.split(':'); el=etree.Element('{%s}%s'%({'p':P,'a':A,'p14':P14}[pre],t))
    for k,v in at.items(): el.set(k,str(v))
    return el
def sub(parent,tag,**at):
    el=E(tag,**at); parent.append(el); return el

def cmd_build(a):
    import numpy as np
    SR=24000
    segs=json.load(open(a.segs)); cues=json.load(open(a.cues)) if a.cues else {}
    d=tempfile.mkdtemp(); unpack(a.deck,d)
    order=slide_order(d)
    ctx=etree.parse(d+'/[Content_Types].xml'); ct=ctx.getroot()
    if not any(x.get('Extension')=='m4a' for x in ct):
        x=etree.Element('{%s}Default'%CT); x.set('Extension','m4a'); x.set('ContentType','audio/mp4'); ct.insert(0,x)
    if not any(x.get('Extension')=='png' for x in ct):
        x=etree.Element('{%s}Default'%CT); x.set('Extension','png'); x.set('ContentType','image/png'); ct.insert(0,x)
    ctx.write(d+'/[Content_Types].xml',xml_declaration=True,encoding='UTF-8',standalone=True)
    from PIL import Image
    Image.new('RGB',(64,64),(128,128,128)).save(d+'/ppt/media/narration_icon.png')
    report=[]; tracks=[]
    for s in segs:
        sl=s['slide']; path=d+'/ppt/'+order[sl-1]
        tree=etree.parse(path); root=tree.getroot()
        ms,gs=groups(root)
        autog=[g for g in gs if g['auto']]; clickg=[g for g in gs if not g['auto']]
        clicks={int(k):v for k,v in s['clicks'].items()}
        nclick=len(clickg)
        if max(clicks,default=0)!=nclick:
            raise SystemExit(f'slide {sl}: notes have {max(clicks,default=0)} clicks but the slide has {nclick} click animations')
        # ---- schedule (seconds from slide start)
        audio=[]; t_end=0.0
        t_auto=max([g['end'] for g in autog],default=0)/1000
        if clicks.get(0):
            t=a.slide_lead
            for sg in clicks[0]: audio.append((t,sg,'시작')); t+=sg['dur']+a.sentence_gap
            t_end=t-a.sentence_gap
        T={}; shifts={}; prev_T=None; prev_anim=0.0
        for k in range(1,nclick+1):
            g=clickg[k-1]; anim=g['end']/1000
            if k==1: tk=max(a.first_click if not clicks.get(0) else t_end+a.gap, (t_auto+a.anim_pad) if autog else 0)
            else: tk=max(t_end+a.gap, prev_T+prev_anim+a.anim_pad)
            T[k]=tk
            segl=clicks.get(k,[]); timed=any(sg['cue']>0 for sg in segl)
            t=tk+(0.3 if timed else min(a.click_lead,max(0.4,anim)))
            cfg=cues.get(str(sl),{}).get(str(k))
            cue_list=[sg['cue'] for sg in segl if sg['cue']>0]
            if timed and not cfg:
                # default: an effect belongs to the first cue at or after it (lead-in effects go with the next sentence)
                ef=sorted(e['start'] for e in g['effects']); b=[0]
                for c in cue_list[:-1]:
                    nxt=[x for x in ef if x>c*1000+50]; b.append(nxt[0] if nxt else int(c*1000)+50)
                cfg={'mode':'first','bounds':b}
            gsh=[]; prev_shift=None; ci=0
            for sg in segl:
                st=t
                if sg['cue']>0:
                    lo=cfg['bounds'][ci]/1000; hi=cfg['bounds'][ci+1]/1000 if ci+1<len(cfg['bounds']) else 1e9
                    anchor=lo if cfg['mode']=='first' else sg['cue']; pre=0.0 if cfg['mode']=='first' else a.cue_pre
                    shift=st-pre-(tk+anchor)
                    if prev_shift is not None: shift=max(shift,prev_shift)
                    st=tk+anchor+shift+pre; gsh.append((lo,hi,shift)); prev_shift=shift; ci+=1
                audio.append((st,sg,f'클릭{k}'+(f" 【{sg['cue']}초】" if sg['cue']>0 else '')))
                t=st+sg['dur']+a.sentence_gap
            if timed and ci!=len(cfg['bounds']): raise SystemExit(f'slide {sl} click {k}: cues.json bounds count != cue sentences')
            shifts[k]=gsh; t_end=t-a.sentence_gap
            prev_T=tk; prev_anim=anim+max([x[2] for x in gsh],default=0)
        last_anim=(prev_T+prev_anim) if nclick else t_auto
        slide_len=max(t_end,last_anim)+a.tail
        # ---- audio track
        n=int(slide_len*SR); buf=np.zeros(n+SR,dtype=np.int16)
        for st,sg,_ in audio:
            w=wave.open(sg['wav']); assert w.getframerate()==SR and w.getsampwidth()==2
            x=np.frombuffer(w.readframes(w.getnframes()),dtype=np.int16); w.close()
            i=int(st*SR); buf[i:i+len(x)]=x[:len(buf)-i]
        buf=buf[:n]; tracks.append(buf)
        tmp=os.path.join(d,'_tmp.wav'); w=wave.open(tmp,'wb'); w.setnchannels(1); w.setsampwidth(2); w.setframerate(SR); w.writeframes(buf.tobytes()); w.close()
        m4a=f'narration_s{sl:02d}.m4a'
        subprocess.run(['ffmpeg','-loglevel','error','-y','-i',tmp,'-c:a','aac','-b:a','128k',d+'/ppt/media/'+m4a],check=True); os.remove(tmp)
        # ---- hidden audio object, off-slide
        spTree=root.find('.//p:cSld/p:spTree',ns)
        spid=max(int(x.get('id')) for x in root.iter('{%s}cNvPr'%P))+1
        rp=rels_path(path)
        if os.path.exists(rp): rt=etree.parse(rp)
        else: rt=etree.ElementTree(etree.Element('{%s}Relationships'%PR))
        rr=rt.getroot(); used={x.get('Id') for x in rr}
        def addrel(typ,target):
            i=1
            while f'rId{i}' in used: i+=1
            rid=f'rId{i}'; used.add(rid)
            el=etree.SubElement(rr,'{%s}Relationship'%PR); el.set('Id',rid); el.set('Type',typ); el.set('Target',target); return rid
        rA=addrel(R+'/audio','../media/'+m4a); rM=addrel('http://schemas.microsoft.com/office/2007/relationships/media','../media/'+m4a)
        rI=addrel(R+'/image','../media/narration_icon.png')
        rt.write(rp,xml_declaration=True,encoding='UTF-8',standalone=True)
        sldW=int(etree.parse(d+'/ppt/presentation.xml').find('.//p:sldSz',ns).get('cx'))
        pic=E('p:pic'); nv=sub(pic,'p:nvPicPr')
        c=sub(nv,'p:cNvPr',id=spid,name='Narration'); h=sub(c,'a:hlinkClick'); h.set('{%s}id'%R,''); h.set('action','ppaction://media')
        sub(sub(nv,'p:cNvPicPr'),'a:picLocks',noChangeAspect='1')
        nvpr=sub(nv,'p:nvPr'); sub(nvpr,'a:audioFile').set('{%s}link'%R,rA)
        ex=sub(sub(nvpr,'p:extLst'),'p:ext',uri='{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}'); sub(ex,'p14:media').set('{%s}embed'%R,rM)
        bf=sub(pic,'p:blipFill'); sub(bf,'a:blip').set('{%s}embed'%R,rI); sub(sub(bf,'a:stretch'),'a:fillRect')
        spp=sub(pic,'p:spPr'); xf=sub(spp,'a:xfrm'); sub(xf,'a:off',x=sldW+152400,y=152400); sub(xf,'a:ext',cx=304800,cy=304800)
        sub(sub(spp,'a:prstGeom',prst='rect'),'a:avLst')
        spTree.append(pic)
        # ---- timing: ONE auto-start group; media first, every other effect "With Previous" + absolute delay.
        # (PowerPoint recomputes "After Previous" chains from its own model and ignores sub-par delays,
        #  so absolute With-Previous delays are the only encoding that plays exactly as scheduled.)
        timing=root.find('p:timing',ns)
        if timing is None:
            timing=E('p:timing'); par=sub(sub(timing,'p:tnLst'),'p:par')
            rc=sub(par,'p:cTn',id=1,dur='indefinite',restart='never',nodeType='tmRoot'); rcl=sub(rc,'p:childTnLst')
            seq=sub(rcl,'p:seq',concurrent='1',nextAc='seek'); ms=sub(seq,'p:cTn',id=2,dur='indefinite',nodeType='mainSeq'); sub(ms,'p:childTnLst')
            for tagn,evt in (('p:prevCondLst','onPrev'),('p:nextCondLst','onNext')):
                sub(sub(sub(seq,tagn),'p:cond',evt=evt,delay=0),'p:tgtEl').append(E('p:sldTgt'))
            ext=root.find('p:extLst',ns)
            (ext.addprevious(timing) if ext is not None else root.append(timing))
        gl=ms.find('p:childTnLst',ns)
        newg=E('p:par'); nc=sub(newg,'p:cTn',id=0,fill='hold'); scl=sub(nc,'p:stCondLst'); sub(scl,'p:cond',delay='indefinite')
        tn=sub(sub(scl,'p:cond',evt='onBegin',delay=0),'p:tn',val=0)
        mc=sub(sub(sub(nc,'p:childTnLst'),'p:par'),'p:cTn',id=0,fill='hold'); sub(sub(mc,'p:stCondLst'),'p:cond',delay=0)
        mcl=sub(mc,'p:childTnLst')
        ec=sub(sub(mcl,'p:par'),'p:cTn',id=0,presetID=1,presetClass='mediacall',presetSubtype=0,fill='hold',nodeType='afterEffect')
        sub(sub(ec,'p:stCondLst'),'p:cond',delay=0)
        cb=sub(sub(sub(ec,'p:childTnLst'),'p:cmd',type='call',cmd='playFrom(0.0)'),'p:cBhvr'); sub(cb,'p:cTn',id=0,dur=int(slide_len*1000),fill='hold')
        sub(sub(cb,'p:tgtEl'),'p:spTgt',spid=spid)
        coll=[]
        def place(g,offset_ms,gsh):
            for e in g['effects']:
                at=e['start']/1000; sh=0
                for lo,hi,v in gsh:
                    if lo-0.001<=at<hi-0.001: sh=v
                absms=offset_ms+e['start']+int(round(sh*1000))
                e['ctn'].set('nodeType','withEffect')
                cond=e['ctn'].find('p:stCondLst/p:cond',ns)
                if cond is None: cond=sub(e['ctn'].find('p:stCondLst',ns) if e['ctn'].find('p:stCondLst',ns) is not None else sub(e['ctn'],'p:stCondLst'),'p:cond')
                cond.set('delay',str(absms)); coll.append((absms,e['par']))
        for g in autog: place(g,0,[])
        for k in range(1,nclick+1): place(clickg[k-1],int(T[k]*1000),shifts[k])
        for _,par in sorted(coll,key=lambda x:x[0]): mcl.append(par)
        for g in gs: gl.remove(g['par'])
        gl.insert(0,newg)
        au=sub(timing.find('p:tnLst/p:par/p:cTn/p:childTnLst',ns),'p:audio'); cm=sub(au,'p:cMediaNode',vol=100000,showWhenStopped=0)
        act=sub(cm,'p:cTn',id=0,fill='hold',display=0); sub(sub(act,'p:stCondLst'),'p:cond',delay='indefinite')
        sub(sub(sub(act,'p:endCondLst'),'p:cond',evt='onStopAudio',delay=0),'p:tgtEl').append(E('p:sldTgt'))
        sub(sub(cm,'p:tgtEl'),'p:spTgt',spid=spid)
        for i,ctn in enumerate(timing.iter('{%s}cTn'%P),1): ctn.set('id',str(i))
        tn.set('val',ms.get('id'))
        for old in root.findall('p:transition',ns): root.remove(old)
        timing.addprevious(E('p:transition',advTm=int(slide_len*1000)))
        tree.write(path,xml_declaration=True,encoding='UTF-8',standalone=True)
        report.append({'slide':sl,'len':round(slide_len,1),'clicks':{k:round(v,2) for k,v in T.items()},
                       'audio':[(round(st,2),lab,sg['text']) for st,sg,lab in audio]})
    pack(d,a.o); shutil.rmtree(d)
    base=os.path.splitext(a.o)[0]
    json.dump(report,open(base+'_timeline.json','w'),ensure_ascii=False,indent=1)
    with open(base+'_timeline.md','w') as f:
        for r in report:
            f.write(f"## 슬라이드 {r['slide']}  ({r['len']}초)\n\n| 시각 | 구간 | 낭독 |\n|---|---|---|\n")
            for st,lab,txt in r['audio']: f.write(f'| {st:.1f}s | {lab} | {txt} |\n')
            f.write('\n')
    allbuf=np.concatenate(tracks); tmp=base+'_review.wav'
    w=wave.open(tmp,'wb'); w.setnchannels(1); w.setsampwidth(2); w.setframerate(SR); w.writeframes(allbuf.tobytes()); w.close()
    subprocess.run(['ffmpeg','-loglevel','error','-y','-i',tmp,'-b:a','96k',base+'_review.mp3'],check=True); os.remove(tmp)
    print(f'{a.o}  total {sum(r["len"] for r in report)/60:.1f} min; timeline + review mp3 written')

def main():
    ap=argparse.ArgumentParser(description=__doc__,formatter_class=argparse.RawDescriptionHelpFormatter); sp=ap.add_subparsers(dest='cmd',required=True)
    x=sp.add_parser('inspect'); x.add_argument('deck'); x.add_argument('-v','--verbose',action='store_true'); x.set_defaults(f=cmd_inspect)
    x=sp.add_parser('parse'); x.add_argument('deck'); x.add_argument('-o',required=True); x.set_defaults(f=cmd_parse)
    for name in ('samples','tts'):
        x=sp.add_parser(name); x.add_argument('--style',default=None); x.add_argument('--rate',default=None); x.add_argument('--lang',default=None,help='SSML language; default: taken from the voice name')
        if name=='samples':
            x.add_argument('--text',required=True); x.add_argument('--voices',required=True); x.add_argument('--outdir',default='samples'); x.add_argument('--prefix',default='sample_'); x.set_defaults(f=cmd_samples)
        else:
            x.add_argument('segs'); x.add_argument('-o',required=True); x.add_argument('--voice',required=True); x.add_argument('--outdir',default='wav'); x.add_argument('--force',action='store_true'); x.set_defaults(f=cmd_tts)
    x=sp.add_parser('build'); x.add_argument('deck'); x.add_argument('segs'); x.add_argument('-o',required=True); x.add_argument('--cues')
    for k,v in (('slide_lead',0.7),('first_click',0.6),('click_lead',1.0),('gap',0.5),('sentence_gap',0.25),('anim_pad',0.2),('tail',1.5),('cue_pre',0.2)):
        x.add_argument('--'+k.replace('_','-'),dest=k,type=float,default=v)
    x.set_defaults(f=cmd_build)
    a=ap.parse_args(); a.f(a)
if __name__=='__main__': main()
