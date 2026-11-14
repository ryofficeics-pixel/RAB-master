/* (PEL3 + TARIF V2: sinkron volume/% pelaksana + mode SAMA/PERSEN/VOLUME; TARIF INDIVIDUAL per pekerjaan+pelaksana via sheet TARIF_PEKERJAAN, menu Tarif) SINKRON FILE ADMIN <- FILE STAFF. Tempel di file ADMIN: Extensions > Apps Script (ganti SELURUH isi SINKRON_ART_PEL3). */
var CFG={"N":141,"LASTR":198,"B0":40,"NT":8,"NP":40,"NSAT":14,"RTB":52,"RTH":53,"RT1":54,"D1":98,"Q":[9,10,11,12,13,16,17,18,19,20,21,22,23,24,25,26,27,28,31,32,33,34,35,36,37,38,39,40,43,44,45,46,47,50,51,52,53,54,55,58,59,62,63,64,65,66,69,70,71,72,73,76,77,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,97,98,99,100,101,104,105,106,109,110,111,114,115,116,119,120,121,122,123,124,125,126,129,130,131,132,133,136,137,138,139,140,143,144,145,146,147,148,149,150,151,154,155,156,157,160,161,162,163,164,167,168,171,172,175,176,177,178,179,180,181,182,183,184,185,186,187,190,191,192,195,196,197]};
function extractId_(v){var m=String(v||'').match(/\/d\/([a-zA-Z0-9_-]+)/);return m?m[1]:String(v||'').trim();}
function baca_(src,sheet,a1){
  var s=src.getSheetByName(sheet);
  if(!s)throw new Error('Sheet tidak ada di file STAFF: '+sheet);
  return s.getRange(a1).getValues();
}
function tulis_(dst,cell,v){
  dst.getRange(cell).offset(0,0,v.length,v[0].length).setValues(v);
}
function sinkron(){
  var ss=SpreadsheetApp.getActive(),su=ss.getSheetByName('SETUP'),ds=ss.getSheetByName('DATA STAF');
  try{
    var id=extractId_(su.getRange('B3').getValue());
    if(!id)throw new Error('ID/URL file STAFF belum diisi di SETUP!B3');
    var st=SpreadsheetApp.openById(id),N=CFG.N,L=CFG.LASTR,n=L-7;
    /* semua blok disejajarkan ke NOMOR BARIS RAB (baris 8..LASTR): baris yang sama = item yang sama */
    var ph=st.getSheetByName('PRODUKTIVITAS HARIAN');if(!ph)throw new Error('Sheet tidak ada di file STAFF: PRODUKTIVITAS HARIAN');
    var v=ph.getRange('B5:J'+(4+N)).getValues(),out=[],i,hit=0;
    for(i=0;i<n;i++)out.push(['','','','','','','','','']);
    v.forEach(function(r){var ref=Number(r[0]);if(ref>=8&&ref<=L){out[ref-8]=r;hit++;}});
    if(hit!==N)throw new Error('Jumlah item di PRODUKTIVITAS HARIAN ('+hit+') tidak sama dengan '+N+'. Apakah ada baris yang disisipkan/dihapus di file STAFF?');
    var bp=st.getSheetByName('BOBOT PEK');if(!bp)throw new Error('Sheet tidak ada di file STAFF: BOBOT PEK');
    /* kolom & baris BOBOT PEK dicari lewat JUDUL kolom + isi uraian (bukan koordinat tetap) */
    var bv=bp.getRange(1,1,Math.max(bp.getLastRow(),L),Math.max(bp.getLastColumn(),21)).getValues(),cU=-1,cC=-1,cS=-1,cV=-1,cN=-1,h0=-1,r,c,t;
    var cP=[-1,-1,-1,-1,-1],pv=[[],[],[],[],[]],cM=-1,mm=[];
    for(r=0;r<Math.min(15,bv.length)&&cU<0;r++)for(c=0;c<bv[r].length;c++){t=String(bv[r][c]).trim().toUpperCase();if(t==='PELAKSANA'&&cU<0){cU=c;h0=r}if(t==='URAIAN PEKERJAAN'&&cC<0)cC=c;if(t==='STATUS ITEM'&&cS<0)cS=c;if(t==='VOLUME REVISI'&&cV<0)cV=c;if(t==='CATATAN STATUS'&&cN<0)cN=c;if(/^(VOLUME PELAKSANA|PELAKSANA) [1-5]/.test(t))cP[Number(t.replace(/^\D+ (\d).*$/,'$1'))-1]=c;if(t==='MODE PEMBAGIAN')cM=c;}
    if(cU<0||cC<0||cS<0||cV<0||cN<0)throw new Error('Judul kolom PELAKSANA / Uraian Pekerjaan / STATUS ITEM / VOLUME REVISI / CATATAN STATUS tidak ditemukan di BOBOT PEK (baris 1-15)');
    var key=function(x){return String(x).replace(/\s+/g,' ').trim();},used={},kk=[],ll=[],ss=[],vv=[],nn=[],miss=[];
    for(i=0;i<n;i++){kk.push(['']);ll.push(['']);ss.push(['']);vv.push(['']);nn.push(['']);pv.forEach(function(a){a.push(['']);});mm.push(['']);}
    v.forEach(function(q){var ref=Number(q[0]);if(!(ref>=8&&ref<=L))return;var it=key(q[3]),rr=ref-1,ok=(rr<bv.length&&key(bv[rr][cC])===it&&!used[rr]);
      if(!ok){var best=-1,d,bd=1e9;for(var z=h0+1;z<bv.length;z++){if(!used[z]&&key(bv[z][cC])===it){d=Math.abs(z-rr);if(d<bd){bd=d;best=z;}}}if(best<0){miss.push(ref);return;}rr=best;}
      used[rr]=1;kk[ref-8]=[bv[rr][cU]];ll[ref-8]=[bv[rr][cC]];ss[ref-8]=[String(bv[rr][cS]).trim().toUpperCase()];vv[ref-8]=[(typeof bv[rr][cV]==='number'&&bv[rr][cV]>0)?bv[rr][cV]:''];nn[ref-8]=[String(bv[rr][cN]||'')];cP.forEach(function(cc,k){var x=cc>=0?bv[rr][cc]:'';pv[k][ref-8]=[(typeof x==='number'&&x>0)?x:''];});mm[ref-8]=[cM>=0?String(bv[rr][cM]||'').trim().toUpperCase():''];});
    if(miss.length)throw new Error('Item tidak ditemukan di BOBOT PEK (baris RAB '+miss.slice(0,10).join(',')+(miss.length>10?'...':'')+'): uraian di BOBOT PEK berbeda dengan PRODUKTIVITAS HARIAN');
    var kal=baca_(st,'KALENDER MINGGU','A'+CFG.B0+':K'+(CFG.B0+CFG.NT-1)),rkn=baca_(st,'REKAP PELAKSANA','B7:B'+(6+CFG.NP));
    /* SEMUA data sudah terbaca & lolos cek; baru menulis (tidak ada data setengah jadi bila gagal) */
    ds.getRange('A8:L'+L).clearContent();
    ds.getRange(8,32,n,6).clearContent();
    ds.getRange(8,1,n,9).setValues(out);
    ds.getRange(8,10,n,1).setValues(ss);
    ds.getRange(8,27,n,1).setValues(vv);
    ds.getRange(8,29,n,1).setValues(nn);
    for(var k2=0;k2<5;k2++)ds.getRange(8,32+k2,n,1).setValues(pv[k2]);
    ds.getRange(8,37,n,1).setValues(mm);
    ds.getRange(8,11,n,1).setValues(kk);
    ds.getRange(8,12,n,1).setValues(ll);
    tulis_(ds,'M5',kal);
    tulis_(ds,'Y5',rkn);
    su.getRange('B5').setValue('OK '+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd-MMM-yyyy HH:mm')+' ('+hit+' item)');
    try{if(manAktif_())manKeAdmin();else su.getRange('B5').setValue(su.getRange('B5').getValue()+' | manual: belum aktif (menu Sinkron > Pasang sinkron manual)');}catch(x){su.getRange('B5').setValue(su.getRange('B5').getValue()+' | manual GAGAL: '+x.message);}
  }catch(e){
    su.getRange('B5').setValue('GAGAL: '+e.message);
    throw e;
  }
}
/* Tarif 2 arah: ketik Upah satuan (kolom H, tabel DETAIL) -> otomatis masuk tabel TARIF (pelaksana x satuan). Tabel TARIF -> kolom H lewat rumus.
   Posisi tabel dicari lewat JUDUL, dan nomor item dibaca dari rumus kolom D baris itu sendiri: aman bila baris disisipkan/dihapus. */
function lokasiTarif_(sh){
  var last=sh.getLastRow(),B=sh.getRange(1,2,last,1).getValues(),H=sh.getRange(1,8,last,1).getValues(),rtb=-1,d1=-1,i;
  for(i=0;i<last;i++){
    if(rtb<0&&String(B[i][0]).indexOf('TARIF UPAH PER PELAKSANA')===0)rtb=i+1;
    if(rtb>0&&d1<0&&i+1>rtb&&String(H[i][0]).indexOf('Upah satuan')===0)d1=i+2;
  }
  return(rtb<0||d1<0)?null:{RTH:rtb+1,RT1:rtb+2,D1:d1};
}
function tarifDariDetail_(sh,ra,rb){
  var ss=SpreadsheetApp.getActive(),k=function(x){return String(x).replace(/\s+/g,' ').trim().toLowerCase();},tst=function(m){ss.toast(m,'Tarif',10);};
  var P=lokasiTarif_(sh);if(!P){tst('Judul tabel TARIF / DETAIL tidak ditemukan (dihapus/diubah?). Tarif tidak disimpan.');return;}
  var a=Math.max(ra,P.D1),b=Math.min(rb,sh.getLastRow());if(a>b)return;
  var n=b-a+1,i,msg=[],ok=0,W=CFG.NSAT+3;
  /* nomor baris RAB tiap item: dari rumus kolom D (='NILAI PROGRES'!F<q>), bukan dari daftar tetap */
  var dF=sh.getRange(a,4,n,1).getFormulas(),Qs=[];
  for(i=0;i<n;i++){var mm=/NILAIPROGRES'!F(\d+)$/.exec(String(dF[i][0]).replace(/\s+/g,''));Qs.push(mm?Number(mm[1]):0);}
  var hfm=function(q){return "='NILAI PROGRES'!N"+q;};
  var restore=function(){for(var j=0;j<n;j++)if(Qs[j])sh.getRange(a+j,8).setFormula(hfm(Qs[j]));};
  var nm=sh.getRange(P.RT1,2,CFG.NP,1).getValues(),A1=sh.getRange(P.RT1,1).getValue(),A2=sh.getRange(P.RT1+CFG.NP-1,1).getValue();
  if(Number(A1)!==1||Number(A2)!==CFG.NP){restore();tst('Susunan tabel TARIF berubah (baris disisipkan/dihapus di dalam tabel). Tarif tidak disimpan. Isi lewat tabel TARIF.');return;}
  var hdr=sh.getRange(P.RTH,3,1,W).getValues()[0],blk=sh.getRange(P.RT1,3,CFG.NP,W).getValues(),
    det=sh.getRange(a,3,n,6).getValues(),cur=sh.getRange(a,8,n,1).getFormulas(),raws=sh.getRange(a,8,n,1).getValues(),
    np=ss.getSheetByName('NILAI PROGRES'),NR_=Math.max(np.getLastRow(),Math.max.apply(null,Qs.concat([1]))),vv=np.getRange(1,22,NR_,2).getValues(),n2=np.getRange(1,30,NR_,1).getValues(),hdrChg=false;
  for(i=0;i<n;i++){
    var q=Qs[i],raw=raws[i][0];
    if(!q){if(raw!==''&&raw!==null)msg.push('baris '+(a+i)+': bukan baris item');continue;}
    if(String(cur[i][0]).replace(/\s+/g,'')===hfm(q))continue; /* rumus masih utuh = tidak diubah */
    var kd=String(vv[q-1][1]).trim(),sd=kd?kd:String(det[i][3]).trim(),pel=k(det[i][0]),sat=k(sd),val,lab=String(det[i][0]).trim()+' / '+sd;
    if(String(n2[q-1][0]).trim()!==''){msg.push(lab+': item dibagi >1 pelaksana, isi tarif lewat tabel TARIF');continue;}
    if(!pel||!sat){msg.push('baris '+(a+i)+': pelaksana/satuan kosong');continue;}
    if(Number(vv[q-1][0])>0){msg.push(lab+': item punya upah khusus (kolom V NILAI PROGRES), kosongkan dulu');continue;}
    if(raw===''||raw===null)val='';
    else{val=Number(raw);if(typeof raw!=='number'||!isFinite(val)||val<0){msg.push(lab+': harus angka >= 0');continue;}}
    var pr=-1,sc=-1,j;
    for(j=0;j<nm.length;j++)if(k(nm[j][0])===pel){pr=j;break;}
    if(pr<0){msg.push(lab+': pelaksana tidak ada di tabel TARIF');continue;}
    for(j=0;j<W;j++)if(k(hdr[j])===sat){sc=j;break;}
    if(sc<0){for(j=0;j<W;j++)if(String(hdr[j]).trim()===''){sc=j;hdr[j]=sd;hdrChg=true;break;}}
    if(sc<0){msg.push(lab+': kolom satuan di tabel TARIF penuh');continue;}
    blk[pr][sc]=val;ok++;
  }
  if(hdrChg)sh.getRange(P.RTH,3,1,W).setValues([hdr]);
  if(ok)sh.getRange(P.RT1,3,CFG.NP,W).setValues(blk);
  restore();
  if(ok||msg.length)tst((ok?ok+' tarif tersimpan ke tabel TARIF (berlaku utk semua item pelaksana+kode/satuan yang sama). ':'')+(msg.length?'Ditolak: '+msg.slice(0,5).join('; '):''));
}
/* ===== MANUAL TAMBAH / KURANG: sinkron 2 arah STAFF <-> ADMIN (blok [MANUAL] ... [AKHIR MANUAL] pada sheet PERUBAHAN) =====
   Kolom dicari lewat judul header (KODE, AREA, URAIAN, STATUS, ALASAN/DESKRIPSI, VOLUME, SAT., T/K, HARGA SATUAN, KETERANGAN),
   jadi aman bila kolom disisip / digeser. Penanda blok ada di satu kolom yang berisi tulisan [MANUAL] (SEK, SUB, [AKHIR MANUAL]).
   Kode (KODE) dihitung rumus, tidak disalin. Harga hanya ada di ADMIN dan diikat ke uraian+satuan. */
var MAN={ST:{sheet:'PERUBAHAN PEKERJAAN'},AD:{sheet:'PERUBAHAN (Rp)'}};
function manCols_(sh,hdr){
  var v=sh.getRange(hdr,1,1,30).getValues()[0],m={ar:0,ur:0,st:0,ds:0,vo:0,sa:0,tk:0,hg:0,ke:0},i,t;
  for(i=0;i<v.length;i++){
    t=String(v[i]).toUpperCase().replace(/^\s+/,'');
    if(t.indexOf('AREA')===0)m.ar=i+1;
    else if(t.indexOf('URAIAN')===0)m.ur=i+1;
    else if(t.indexOf('STATUS')===0)m.st=i+1;
    else if(t.indexOf('ALASAN')===0||t.indexOf('DESKRIPSI')===0)m.ds=i+1;
    else if(t.indexOf('VOLUME')===0)m.vo=i+1;
    else if(t.indexOf('SAT')===0)m.sa=i+1;
    else if(t==='T/K')m.tk=i+1;
    else if(t.indexOf('HARGA SATUAN')===0)m.hg=i+1;
    else if(t.indexOf('KETERANGAN')===0)m.ke=i+1;
  }
  var miss=['ur','st','ds','vo','sa','tk','ke'].filter(function(k){return !m[k]});
  if(miss.length)throw new Error('Header manual tidak lengkap di '+sh.getName()+' (baris '+hdr+'): '+miss.join(', '));
  m.last=Math.max(m.ke,m.hg,m.ar,m.ur,m.st,m.ds,m.vo,m.sa,m.tk)+3;
  return m;
}
function manSecs_(sh){
  var last=sh.getLastRow(),W=Math.max(sh.getLastColumn(),20),V=sh.getRange(1,1,last,W).getValues(),
      mc=0,i,j,t,s=-1,e=-1,secs=[],cur=null;
  for(i=0;i<last&&!mc;i++)for(j=0;j<W;j++)if(String(V[i][j]).trim()==='[MANUAL]'){mc=j+1;s=i+1;break;}
  if(!mc)throw new Error('Penanda [MANUAL] tidak ditemukan di sheet '+sh.getName()+'. Blok dihapus / diubah?');
  for(i=s;i<last;i++){
    t=String(V[i][mc-1]).trim();
    if(t==='[AKHIR MANUAL]'){e=i+1;break;}
    if(t==='SEK'){cur={h:i+1,sub:0,t:''};secs.push(cur);for(j=0;j<W;j++){var q=String(V[i][j]).trim();if(q.length>2){cur.t=manKey_(q);break;}}}
    else if(t==='SUB'&&cur&&!cur.sub)cur.sub=i+1;
  }
  if(e<0)throw new Error('Penanda [AKHIR MANUAL] tidak ditemukan di sheet '+sh.getName());
  secs.forEach(function(x){if(!x.sub||x.sub-x.h<2)throw new Error('Seksi manual rusak di '+sh.getName()+' (baris '+x.h+'): minimal 1 baris data antara judul seksi dan Subtotal.');});
  var hdr=s+1;for(i=0;i<s;i++)if(String(V[i][mc-1]).trim()==='[HDR]'){hdr=i+1;break;}
  return {s:s,e:e,mc:mc,secs:secs,hdr:hdr};
}
function manKey_(x){return String(x==null?'':x).replace(/ +/g,' ').trim().toLowerCase();}
function manRead_(sh,C,info){
  return info.secs.map(function(x){
    var n=x.sub-x.h-1,v=sh.getRange(x.h+1,1,n,C.last).getValues();
    return v.map(function(r){
      var g=function(k){return C[k]?r[C[k]-1]:'';};
      return {ar:g('ar'),tk:g('tk'),ur:g('ur'),st:g('st'),ds:g('ds'),vo:g('vo'),sa:g('sa'),ke:g('ke'),hg:g('hg')};
    });
  });
}
function manAlign_(sh,x,n,C){
  var cur=x.sub-x.h-1,nc=C.last;
  while(cur<n){
    sh.insertRowBefore(x.sub-1);
    sh.getRange(x.sub,1,1,nc).copyTo(sh.getRange(x.sub-1,1,1,nc));
    x.sub++;cur++;
  }
  while(cur>n&&cur>1){sh.deleteRow(x.sub-1);x.sub--;cur--;}
}
/* pasangkan seksi antar file lewat judul (bukan nomor), supaya beda jumlah seksi tidak merusak sinkron */
function manPairs_(src,dst){
  var out={secs:[],idx:[],miss:[]};
  dst.secs.forEach(function(x){var i=src.secs.findIndex(function(y){return y.t===x.t;});if(i>=0){out.secs.push(x);out.idx.push(i);}else out.miss.push(x.t);});
  return out;
}
function manWrite_(sh,C,info,data,prices){
  var k,i,col;
  for(k=info.secs.length-1;k>=0;k--){
    var x=info.secs[k],rows=data[k]||[];
    manAlign_(sh,x,Math.max(1,rows.length),C);
    var n=x.sub-x.h-1,r0=x.h+1;
    col=function(f){var o=[];for(i=0;i<n;i++)o.push([rows[i]&&rows[i][f]!=null?rows[i][f]:'']);return o;};
    ['ar','tk','ur','st','ds','vo','sa','ke'].forEach(function(f){if(C[f])sh.getRange(r0,C[f],n,1).setValues(col(f));});
    if(C.hg&&prices){var hv=[];for(i=0;i<n;i++){var q=rows[i]?prices[manKey_(rows[i].ur)+'|'+manKey_(rows[i].sa)]:'';hv.push([q==null?'':q]);}sh.getRange(r0,C.hg,n,1).setValues(hv);}
  }
}
function manStaff_(){
  var ad=SpreadsheetApp.getActive(),su=ad.getSheetByName('SETUP'),id=extractId_(su.getRange('B3').getValue());
  if(!id)throw new Error('ID/URL file STAFF belum diisi di SETUP!B3');
  var st=SpreadsheetApp.openById(id),sh=st.getSheetByName(MAN.ST.sheet);
  if(!sh)throw new Error('Sheet tidak ada di file STAFF: '+MAN.ST.sheet);
  return {ad:ad,st:st,shS:sh,shA:ad.getSheetByName(MAN.AD.sheet)};
}
/* STAFF -> ADMIN (baris + isi dari STAFF; harga ADMIN dipertahankan lewat uraian+sat) */
function manKeAdmin(){
  var o=manStaff_(),iS=manSecs_(o.shS),iA=manSecs_(o.shA),cS=manCols_(o.shS,iS.hdr),cA=manCols_(o.shA,iA.hdr);
  var dS=manRead_(o.shS,cS,iS),dA=manRead_(o.shA,cA,iA),prices={},p=manPairs_(iS,iA);
  dA.forEach(function(sec){sec.forEach(function(r){if(manKey_(r.ur)&&r.hg!==''&&r.hg!=null)prices[manKey_(r.ur)+'|'+manKey_(r.sa)]=r.hg;});});
  PropertiesService.getDocumentProperties().setProperty('MAN_TS',String(Date.now()));
  manWrite_(o.shA,cA,{secs:p.secs},p.idx.map(function(i){return dS[i];}),prices);
  PropertiesService.getDocumentProperties().setProperty('MAN_TS',String(Date.now()));
  var msg='Manual STAFF -> ADMIN selesai'+(p.miss.length?'. Seksi di ADMIN tidak ada: '+p.miss.join(' | '):'');
  SpreadsheetApp.getActive().toast(msg,'Sinkron',8);
}
/* ADMIN -> STAFF (tanpa harga) */
function manKeStaff(){
  var o=manStaff_(),iS=manSecs_(o.shS),iA=manSecs_(o.shA),cS=manCols_(o.shS,iS.hdr),cA=manCols_(o.shA,iA.hdr);
  var dA=manRead_(o.shA,cA,iA),p=manPairs_(iA,iS);
  manWrite_(o.shS,cS,{secs:p.secs},p.idx.map(function(i){return dA[i];}),null);
  var msg='Manual ADMIN -> STAFF selesai'+(p.miss.length?'. Seksi di STAFF tidak ada di ADMIN (tidak diubah): '+p.miss.join(' | '):'');
  SpreadsheetApp.getActive().toast(msg,'Sinkron',8);
}
function manAktif_(){return ScriptApp.getProjectTriggers().some(function(t){return t.getHandlerFunction()==='onEditManual';});}
function onEditManual(e){
  try{
    if(!e||!e.range)return;
    var sh=e.range.getSheet();if(sh.getName()!==MAN.AD.sheet)return;
    var info=manSecs_(sh),C=manCols_(sh,info.hdr),r1=e.range.getRow(),r2=e.range.getLastRow();
    if(r2<info.s||r1>info.e)return;
    var c1=e.range.getColumn(),c2=e.range.getLastColumn(),hit=false,c,d=[C.ar,C.ur,C.st,C.ds,C.vo,C.sa,C.tk,C.ke];
    for(c=c1;c<=c2;c++)if(d.indexOf(c)>=0)hit=true;
    if(hit)manKeStaff();
  }catch(x){SpreadsheetApp.getActive().toast('Sinkron manual gagal: '+x.message,'Sinkron',10);}
}
function onChangeManual(e){
  try{
    if(!e||(e.changeType!=='INSERT_ROW'&&e.changeType!=='REMOVE_ROW'))return;
    var ts=Number(PropertiesService.getDocumentProperties().getProperty('MAN_TS')||0);
    if(Date.now()-ts<30000)return;
    var sh=SpreadsheetApp.getActiveSheet();if(!sh||sh.getName()!==MAN.AD.sheet)return;
    manKeStaff();
  }catch(x){SpreadsheetApp.getActive().toast('Sinkron manual gagal: '+x.message,'Sinkron',10);}
}
function pasangManual(){
  ScriptApp.getProjectTriggers().forEach(function(t){var h=t.getHandlerFunction();if(h==='onEditManual'||h==='onChangeManual')ScriptApp.deleteTrigger(t);});
  var ss=SpreadsheetApp.getActive();
  ScriptApp.newTrigger('onEditManual').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('onChangeManual').forSpreadsheet(ss).onChange().create();
  ss.toast('Sinkron manual 2 arah aktif (ADMIN -> STAFF langsung; STAFF -> ADMIN saat Sinkron).','Sinkron',8);
}
function matikanOtomatis(){
  ScriptApp.getProjectTriggers().forEach(function(t){if(t.getHandlerFunction()==='sinkron')ScriptApp.deleteTrigger(t);});
}
function pasangOtomatis(){
  matikanOtomatis();
  ScriptApp.newTrigger('sinkron').timeBased().everyMinutes(15).create();
}

/* ===================== TARIF INDIVIDUAL PER PEKERJAAN + PELAKSANA (V2) =====================
   KUNCI = ID PEKERJAAN | PELAKSANA | KODE TARIF | SATUAN. ID pekerjaan = kolom Q RAB UPDATE.
   Alur: menu Tarif > A. Persiapan > B. Validasi > C. AKTIFKAN. Rollback tersedia. onEdit/onOpen di bawah menggantikan versi lama. */
var TV2={SH:'TARIF_PEKERJAAN',BKI:'BACKUP_TARIF_ITEM',BKG:'BACKUP_TARIF_GLOBAL',PROP:'TARIF_V2',RABID:17,CID:49,CST:50,CNOTE:'Tidak dipakai lagi setelah tarif individual aktif. Isi tarif di sheet TARIF_PEKERJAAN.'};
/* NILAI PROGRES: B uraian, D sat, M pelaksana, S status, V upah khusus (lama), W kode tarif, O ongkos item,
   AC:AG nama pelaksana 1-5 (col 29-33), AM:AQ tarif satuan pelaksana 1-5 (col 39-43), AR:AV ongkos (col 44-48),
   AW(49)=ID PEKERJAAN (baru), AX(50)=STATUS TARIF (baru) */

/* ---------- util ---------- */
function tarifAktif_(){return PropertiesService.getDocumentProperties().getProperty(TV2.PROP)==='1';}
function tk_(x){return String(x==null?'':x).replace(/\s+/g,' ').trim().toLowerCase();}
function tkey_(id,pel,kode,sat){return tk_(id)+'|'+tk_(pel)+'|'+tk_(kode)+'|'+tk_(sat);}
function colL_(n){var s='',m;while(n>0){m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s;}
function tarifKunci_(fn){
  var lk=LockService.getDocumentLock();
  if(!lk.tryLock(30000))throw new Error('Spreadsheet sedang dipakai proses lain. Coba lagi sebentar lagi.');
  try{return fn();}finally{lk.releaseLock();}
}
function tarifB50_(rp){
  var v=rp.getRange(1,2,Math.min(rp.getLastRow(),120),1).getValues();
  for(var i=0;i<v.length;i++)if(String(v[i][0]).indexOf('Item bertugas tapi tarif belum diisi')===0)return i+1;
  return 0;
}

/* ---------- baca item (satu-satunya tempat membaca NILAI PROGRES) ---------- */
function tarifItems_(){
  var ss=SpreadsheetApp.getActive(),np=ss.getSheetByName('NILAI PROGRES'),ds=ss.getSheetByName('DATA STAF'),L=CFG.LASTR,n=L-7;
  if(!np||!ds)throw new Error('Sheet NILAI PROGRES / DATA STAF tidak ditemukan');
  var A=np.getRange(8,1,n,1).getFormulas(),V=np.getRange(8,1,n,50).getValues(),R=ds.getRange(8,1,n,1).getValues(),out=[],i,k,m,v,names;
  for(i=0;i<n;i++){
    m=/'RABUPDATE'!B(\d+)$/.exec(String(A[i][0]).replace(/\s+/g,''));
    if(!m||Number(R[i][0])!==i+8)continue;               /* bukan baris item */
    v=V[i];names=[];
    for(k=0;k<5;k++)names.push(String(v[28+k]==null?'':v[28+k]).trim());
    out.push({r:i+8,x:Number(m[1]),uraian:String(v[1]==null?'':v[1]),sat:String(v[3]==null?'':v[3]).trim(),
      status:String(v[18]==null?'':v[18]).trim().toUpperCase(),vk:Number(v[21])||0,kode:String(v[22]==null?'':v[22]).trim(),
      names:names,am:[38,39,40,41,42].map(function(z){return Number(v[z])||0;}),
      ong:[43,44,45,46,47].map(function(z){return Number(v[z])||0;}),oc:Number(v[14])||0,id:String(v[48]==null?'':v[48]).trim()});
  }
  return out;
}

/* ---------- sheet TARIF_PEKERJAAN ---------- */
function tarifSheet_(create){
  var ss=SpreadsheetApp.getActive(),sh=ss.getSheetByName(TV2.SH);
  if(sh||!create)return sh;
  sh=ss.insertSheet(TV2.SH);
  var H=['KUNCI (otomatis)','ID_PEKERJAAN','URAIAN PEKERJAAN','PELAKSANA','KODE_TARIF','SATUAN','TARIF (Rp/sat)','STATUS','SUMBER','DIUBAH','CATATAN'];
  sh.getRange(1,1,1,11).setValues([H]).setFontWeight('bold').setBackground('#1F3B73').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  [140,100,280,130,90,70,120,110,130,120,300].forEach(function(w,i){sh.setColumnWidth(i+1,w);});
  sh.getRange('A2:A5000').setFontColor('#999999');
  sh.getRange('G2:G5000').setNumberFormat('#,##0.##').setBackground('#FFF9E5')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).setHelpText('Isi angka >= 0 (Rp per satuan). Kosong = BELUM DIISI.').build());
  sh.getRange('J2:J5000').setNumberFormat('dd-mmm-yyyy hh:mm');
  sh.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$H2="BELUM DIISI"').setBackground('#FCE4D6').setRanges([sh.getRange('A2:K5000')]).build()]);
  return sh;
}
function tarifBaca_(sh){
  var last=sh?sh.getLastRow():0,rows=[],idx={},dup=[],i,v,key,o;
  if(last>=2){
    v=sh.getRange(2,2,last-1,10).getValues();                         /* B..K */
    for(i=0;i<v.length;i++){
      if(String(v[i][0]).trim()===''&&String(v[i][2]).trim()==='')continue;
      key=tkey_(v[i][0],v[i][2],v[i][3],v[i][4]);
      o={r:i+2,id:v[i][0],uraian:v[i][1],pel:v[i][2],kode:v[i][3],sat:v[i][4],tarif:v[i][5],cat:v[i][9],key:key};
      rows.push(o);
      if(idx[key])dup.push(key);else idx[key]=o;
    }
  }
  return {rows:rows,idx:idx,dup:dup,last:last};
}
function tarifTulis_(sh,list){            /* list: [id,uraian,pel,kode,sat,tarif,sumber,catatan] */
  if(!list.length)return;
  var r0=Math.max(sh.getLastRow(),1)+1,now=new Date(),out=list.map(function(x,i){
    var r=r0+i;
    return ['=LOWER(TRIM(B'+r+'))&"|"&LOWER(TRIM(D'+r+'))&"|"&LOWER(TRIM(E'+r+'))&"|"&LOWER(TRIM(F'+r+'))',
      x[0],x[1],x[2],x[3],x[4],x[5],
      '=IF(G'+r+'="","BELUM DIISI",IF(ISNUMBER(G'+r+'),IF(G'+r+'>=0,"TERISI","TARIF NEGATIF"),"BUKAN ANGKA"))',x[6]||'',now,x[7]||''];
  });
  sh.getRange(r0,1,out.length,11).setValues(out);
}

/* ---------- LANGKAH 1: BACKUP ---------- */
function tarifBackup_(){
  if(tarifAktif_())return 'dilewati (tarif individual sudah aktif; backup asli dipertahankan)';
  var ss=SpreadsheetApp.getActive(),np=ss.getSheetByName('NILAI PROGRES'),rp=ss.getSheetByName('REKAP PELAKSANA (Rp)'),L=CFG.LASTR,n=L-7,i,k;
  var V=np.getRange(8,1,n,50).getValues(),F=np.getRange(8,39,n,5).getFormulas();
  var bi=ss.getSheetByName(TV2.BKI)||ss.insertSheet(TV2.BKI);bi.clear();
  var head=['Baris NP','AM','AN','AO','AP','AQ','F_AM','F_AN','F_AO','F_AP','F_AQ','V upah khusus','W kode','D satuan','M pelaksana','S status'],rows=[],sumO=0,sumOng=0;
  for(i=0;i<n;i++){
    var v=V[i];sumO+=Number(v[14])||0;for(k=43;k<=47;k++)sumOng+=Number(v[k])||0;
    rows.push([i+8,v[38],v[39],v[40],v[41],v[42],F[i][0],F[i][1],F[i][2],F[i][3],F[i][4],v[21],v[22],v[3],v[12],v[18]]);
  }
  bi.getRange(1,1,1,16).setValues([head]).setFontWeight('bold');
  bi.getRange(2,7,n,5).setNumberFormat('@');                          /* rumus disimpan sebagai TEKS */
  bi.getRange(2,1,n,16).setValues(rows);
  var b50=tarifB50_(rp),fb=b50?rp.getRange(b50,2).getFormula():'';
  bi.getRange(1,18,6,1).setValues([['B50_BARIS'],['B50_RUMUS'],['SUM_O'],['SUM_ONGKOS'],['ITEM'],['WAKTU']]);
  bi.getRange(2,19).setNumberFormat('@');
  bi.getRange(1,19,6,1).setValues([[b50],[fb],[sumO],[sumOng],[n],[new Date()]]);
  var P=lokasiTarif_(rp)||{RTH:CFG.RTH},bg=ss.getSheetByName(TV2.BKG)||ss.insertSheet(TV2.BKG);bg.clear();
  var g=rp.getRange(P.RTH,2,1+CFG.NP,CFG.NSAT+4).getValues();
  bg.getRange(1,1,g.length,g[0].length).setValues(g);
  bg.getRange(g.length+2,1).setValue('Backup tabel TARIF lama (pelaksana x satuan) - '+new Date());
  return 'OK: '+n+' baris + tabel tarif lama';
}

/* ---------- LANGKAH 2: SIAPKAN (sheet, ID pekerjaan, kolom bantu) ---------- */
function tarifSiapkan_(){
  var ss=SpreadsheetApp.getActive(),np=ss.getSheetByName('NILAI PROGRES'),ru=ss.getSheetByName('RAB UPDATE'),L=CFG.LASTR,n=L-7;
  tarifSheet_(true);
  var it=tarifItems_(),len=Math.max(ru.getLastRow(),L+5),ids=ru.getRange(1,TV2.RABID,len,1).getValues(),seen={},nb=0,chg=false;
  it.forEach(function(x){
    var cur=String(ids[x.x-1][0]).trim();
    if(!cur||seen[cur]){cur='PK-'+Utilities.getUuid().replace(/-/g,'').slice(0,8).toUpperCase();ids[x.x-1][0]=cur;chg=true;nb++;}
    seen[cur]=1;
  });
  ru.getRange(6,TV2.RABID).setValue('ID PEKERJAAN (jangan diubah)').setFontWeight('bold');
  if(chg)ru.getRange(1,TV2.RABID,len,1).setValues(ids);
  ru.getRange(7,TV2.RABID,len-6,1).setFontColor('#888888');ru.setColumnWidth(TV2.RABID,110);
  var fa=[];for(var i=0;i<n;i++)fa.push(['']);
  it.forEach(function(x){fa[x.r-8]=["='RAB UPDATE'!"+colL_(TV2.RABID)+x.x];});
  np.getRange(6,TV2.CID).setValue('ID PEKERJAAN (dari RAB UPDATE kolom Q)');
  np.getRange(6,TV2.CST).setValue('Status tarif individual');
  np.getRange(8,TV2.CID,n,1).setFormulas(fa);
  SpreadsheetApp.flush();
  return nb;
}

/* ---------- LANGKAH 3: MIGRASI tarif lama -> TARIF_PEKERJAAN (dryRun=true: hanya hitung) ---------- */
function tarifMigrasi_(dry){
  if(tarifAktif_())throw new Error('Tarif individual sudah aktif. Migrasi dari tarif lama tidak diperlukan.');
  var sh=tarifSheet_(false),tb=sh?tarifBaca_(sh):{idx:{}},add=[],seen={},c={ada:0,global:0,v:0,cek:0,kosong:0,tanpaId:[]};
  tarifItems_().forEach(function(x){
    if(x.status==='BATAL')return;
    if(!x.id){c.tanpaId.push(x.r);return;}
    var nn=x.names.filter(Boolean).length;
    x.names.forEach(function(p,k){
      if(!p)return;
      var key=tkey_(x.id,p,x.kode,x.sat),t=x.am[k],sum='',cat='';
      if(tb.idx[key]||seen[key]){c.ada++;return;}
      seen[key]=1;
      if(t>0){
        if(x.vk>0){sum='MIGRASI-V';c.v++;if(nn>1){cat='PERLU DICEK: upah khusus item (kolom V) lama berlaku SAMA untuk semua pelaksana';c.cek++;}}
        else{sum='MIGRASI-GLOBAL';c.global++;}
      }else{t='';cat='BELUM ADA TARIF LAMA (tidak diisi otomatis)';c.kosong++;}
      add.push([x.id,x.uraian,p,x.kode,x.sat,t,sum,cat]);
    });
  });
  if(!dry){tarifTulis_(sh,add);SpreadsheetApp.flush();}
  c.baru=add.length;
  return c;
}

/* ---------- VALIDASI (hanya baca) ---------- */
function tarifValidasi_(){
  var sh=tarifSheet_(false);if(!sh)return {fatal:'Sheet TARIF_PEKERJAAN belum ada. Jalankan "A. Persiapan" dulu.'};
  var it=tarifItems_(),tb=tarifBaca_(sh),r={mis:[],kosong:0,tanpaId:[],dupId:[],dupKey:tb.dup,rek:[],orphan:0,sumOng:0,items:0},cnt={},exp={};
  it.forEach(function(x){if(x.id)cnt[x.id]=(cnt[x.id]||0)+1;});
  Object.keys(cnt).forEach(function(k){if(cnt[k]>1)r.dupId.push(k);});
  it.forEach(function(x){
    if(x.status==='BATAL')return;
    r.items++;
    if(!x.id){r.tanpaId.push(x.r);return;}
    x.names.forEach(function(p,k){
      if(!p)return;
      var key=tkey_(x.id,p,x.kode,x.sat),row=tb.idx[key],t=(row&&typeof row.tarif==='number'&&row.tarif>=0)?row.tarif:0;
      exp[key]=1;
      if(!row||row.tarif==='')r.kosong++;
      if(Math.abs(t-x.am[k])>0.005)r.mis.push('baris '+x.r+' '+p+': tabel '+t+' vs sheet '+x.am[k]);
    });
    if(Math.abs(x.ong.reduce(function(a,b){return a+b;},0)-x.oc)>1)r.rek.push(x.r);
    r.sumOng+=x.ong.reduce(function(a,b){return a+b;},0);
  });
  tb.rows.forEach(function(w){if(!exp[w.key])r.orphan++;});
  r.ok=!r.mis.length&&!r.tanpaId.length&&!r.dupId.length&&!r.dupKey.length&&!r.rek.length;
  return r;
}
function tarifLaporan_(r){
  if(r.fatal)return r.fatal;
  return (r.ok?'LOLOS':'GAGAL')+'\n- Item aktif: '+r.items+'\n- Tarif BELUM DIISI: '+r.kosong+
    '\n- Selisih tabel vs sheet: '+r.mis.length+(r.mis.length?' ('+r.mis.slice(0,5).join('; ')+')':'')+
    '\n- Item tanpa ID: '+r.tanpaId.length+'\n- ID ganda: '+r.dupId.length+'\n- Kunci tarif ganda: '+r.dupKey.length+
    '\n- Rekonsiliasi (jumlah upah individual <> total upah item) gagal: '+r.rek.length+(r.rek.length?' (baris '+r.rek.slice(0,8).join(',')+')':'')+
    '\n- Baris tarif tidak terpakai: '+r.orphan;
}

/* ---------- AKTIFKAN: pasang rumus baru AM:AQ, AX, B50 ---------- */
function tarifKeyExpr_(r,k){
  return 'LOWER(TRIM($'+colL_(TV2.CID)+r+'))&"|"&LOWER(TRIM('+colL_(29+k)+r+'))&"|"&LOWER(TRIM($W'+r+'&""))&"|"&LOWER(TRIM($D'+r+'))';
}
function tarifFAM_(r,k){
  var nm=colL_(29+k)+r;
  return '=IF(OR('+nm+'="",UPPER(TRIM($S'+r+'))="BATAL"),0,N(IFERROR(INDEX('+TV2.SH+'!$G:$G,MATCH('+tarifKeyExpr_(r,k)+','+TV2.SH+'!$A:$A,0)),0)))';
}
function tarifFAX_(r){
  var mk=[],k,tj;
  for(k=0;k<5;k++){var nm=colL_(29+k)+r;
    mk.push('IF('+nm+'="","",IF(IFERROR(INDEX('+TV2.SH+'!$H:$H,MATCH('+tarifKeyExpr_(r,k)+','+TV2.SH+'!$A:$A,0)),"")="TERISI","",'+nm+'))');}
  tj='TEXTJOIN(", ",TRUE,'+mk.join(',')+')';
  return '=IF(OR($M'+r+'="",UPPER(TRIM($S'+r+'))="BATAL"),"",IF(TRIM($'+colL_(TV2.CID)+r+')="","ID PEKERJAAN KOSONG",IF('+tj+'="","OK","BELUM DIISI: "&'+tj+')))';
}
function tarifAktifkan_(){
  var ss=SpreadsheetApp.getActive(),np=ss.getSheetByName('NILAI PROGRES'),rp=ss.getSheetByName('REKAP PELAKSANA (Rp)'),L=CFG.LASTR,n=L-7;
  if(tarifAktif_())return {msg:'Tarif individual sudah AKTIF.'};
  if(!ss.getSheetByName(TV2.BKI)||!ss.getSheetByName(TV2.BKG))return {msg:'Backup belum ada. Jalankan "A. Persiapan" dulu.'};
  var pre=tarifValidasi_();
  if(pre.fatal||!pre.ok)return {msg:'Validasi GAGAL, tidak ada yang diubah.\n'+tarifLaporan_(pre)};
  var F=np.getRange(8,39,n,5).getFormulas(),items=tarifItems_(),i,k,fx=[];
  for(i=0;i<n;i++){var rr=[];for(k=0;k<5;k++)rr.push(F[i][k]!==''?tarifFAM_(i+8,k):'');F[i]=rr;fx.push(['']);}
  items.forEach(function(x){fx[x.r-8]=[tarifFAX_(x.r)];});
  PropertiesService.getDocumentProperties().setProperty(TV2.PROP,'1');
  try{
    np.getRange(8,39,n,5).setFormulas(F);
    np.getRange(8,TV2.CST,n,1).setFormulas(fx);
    var b50=tarifB50_(rp);
    if(b50)rp.getRange(b50,2).setFormula('="Item bertugas tapi tarif belum diisi: "&(COUNTIF(\'NILAI PROGRES\'!$AX$8:$AX$'+L+',"BELUM DIISI*")+COUNTIF(\'NILAI PROGRES\'!$AX$8:$AX$'+L+',"ID PEKERJAAN*"))');
    np.getRange(6,22).setNote(TV2.CNOTE);
    SpreadsheetApp.flush();
    var post=tarifValidasi_();
    if(!post.ok||Math.abs(post.sumOng-pre.sumOng)>1)throw new Error('Cek sesudah pemasangan gagal ('+tarifLaporan_(post)+'; total upah sebelum '+pre.sumOng+' sesudah '+post.sumOng+')');
  }catch(e){
    tarifRollback_();
    return {msg:'Pemasangan GAGAL dan sudah di-ROLLBACK otomatis: '+e.message};
  }
  return {msg:'AKTIF. Total upah individual sebelum = sesudah ('+Math.round(pre.sumOng)+').\nTarif BELUM DIISI: '+pre.kosong+' (isi di sheet TARIF_PEKERJAAN).'};
}

/* ---------- ROLLBACK ---------- */
function tarifRollback_(){
  var ss=SpreadsheetApp.getActive(),np=ss.getSheetByName('NILAI PROGRES'),rp=ss.getSheetByName('REKAP PELAKSANA (Rp)'),bi=ss.getSheetByName(TV2.BKI),L=CFG.LASTR,n=L-7;
  if(!bi)throw new Error('Sheet '+TV2.BKI+' tidak ada: tidak bisa rollback otomatis');
  var d=bi.getRange(2,1,n,16).getValues(),F=d.map(function(r){return [r[6],r[7],r[8],r[9],r[10]].map(function(f){return String(f||'');});});
  np.getRange(8,39,n,5).setFormulas(F);                       /* rumus AM:AQ asli (teks '=...' -> rumus) */
  np.getRange(8,TV2.CST,n,1).clearContent();
  var info=bi.getRange(1,19,2,1).getValues(),b50=Number(info[0][0]);
  if(b50&&info[1][0])rp.getRange(b50,2).setFormula(String(info[1][0]));
  np.getRange(6,22).clearNote();
  PropertiesService.getDocumentProperties().setProperty(TV2.PROP,'0');
  SpreadsheetApp.flush();
}

/* ---------- PERBARUI DAFTAR TARIF (tambah baris untuk pekerjaan/pelaksana baru; tandai yang tidak terpakai) ---------- */
function tarifBangun_(){
  var sh=tarifSheet_(true),tb=tarifBaca_(sh),exp={},add=[],seen={},orph=0;
  tarifItems_().forEach(function(x){
    if(!x.id)return;
    x.names.forEach(function(p){
      if(!p)return;
      var key=tkey_(x.id,p,x.kode,x.sat);exp[key]=1;
      if(x.status==='BATAL'||tb.idx[key]||seen[key])return;
      seen[key]=1;add.push([x.id,x.uraian,p,x.kode,x.sat,'','','BARU: isi tarif']);
    });
  });
  if(tb.rows.length){
    var ks=sh.getRange(2,11,tb.last-1,1).getValues();
    tb.rows.forEach(function(w){
      var cur=String(ks[w.r-2][0]);
      if(!exp[w.key]){orph++;if(cur===''||cur==='BARU: isi tarif')ks[w.r-2][0]='TIDAK TERPAKAI';}
      else if(cur==='TIDAK TERPAKAI')ks[w.r-2][0]='';
    });
    sh.getRange(2,11,tb.last-1,1).setValues(ks);
  }
  tarifTulis_(sh,add);
  return {baru:add.length,orphan:orph};
}

/* ---------- GANTI NAMA PELAKSANA (pindahkan tarif ke nama baru; ganti juga di BOBOT PEK file STAFF) ---------- */
function tarifGantiNama_(lama,baru){
  var sh=tarifSheet_(false);if(!sh)throw new Error('Sheet TARIF_PEKERJAAN belum ada');
  lama=String(lama||'').trim();baru=String(baru||'').trim();
  if(!lama||!baru)throw new Error('Nama lama dan baru harus diisi');
  var tb=tarifBaca_(sh),rows=tb.rows.filter(function(w){return tk_(w.pel)===tk_(lama);});
  if(!rows.length)return 0;
  rows.forEach(function(w){if(tb.idx[tkey_(w.id,baru,w.kode,w.sat)]&&tk_(lama)!==tk_(baru))throw new Error('Nama baru sudah punya tarif untuk pekerjaan '+w.uraian+'. Tidak ada yang diubah.');});
  rows.forEach(function(w){sh.getRange(w.r,4).setValue(baru);sh.getRange(w.r,10).setValue(new Date());});
  return rows.length;
}

/* ---------- ketik Upah satuan di DETAIL (REKAP PELAKSANA (Rp), kolom H) saat tarif individual AKTIF ---------- */
function tarifDariDetailV2_(sh,ra,rb){
  var ss=SpreadsheetApp.getActive(),tst=function(m){ss.toast(m,'Tarif',10);},P=lokasiTarif_(sh);
  if(!P){tst('Judul tabel DETAIL tidak ditemukan. Tarif tidak disimpan. Isi lewat sheet TARIF_PEKERJAAN.');return;}
  var a=Math.max(ra,P.D1),b=Math.min(rb,sh.getLastRow());if(a>b)return;
  var n=b-a+1,i,msg=[],ok=0,dF=sh.getRange(a,4,n,1).getFormulas(),raws=sh.getRange(a,8,n,1).getValues(),cur=sh.getRange(a,8,n,1).getFormulas(),Qs=[];
  for(i=0;i<n;i++){var mm=/NILAIPROGRES'!F(\d+)$/.exec(String(dF[i][0]).replace(/\s+/g,''));Qs.push(mm?Number(mm[1]):0);}
  var hfm=function(q){return "='NILAI PROGRES'!N"+q;},restore=function(){for(var j=0;j<n;j++)if(Qs[j])sh.getRange(a+j,8).setFormula(hfm(Qs[j]));};
  var byRow={};tarifItems_().forEach(function(x){byRow[x.r]=x;});
  var tsh=tarifSheet_(false);if(!tsh){restore();tst('Sheet TARIF_PEKERJAAN tidak ada.');return;}
  var tb=tarifBaca_(tsh),add=[];
  for(i=0;i<n;i++){
    var q=Qs[i],raw=raws[i][0];
    if(!q){if(raw!==''&&raw!==null)msg.push('baris '+(a+i)+': bukan baris item');continue;}
    if(String(cur[i][0]).replace(/\s+/g,'')===hfm(q))continue;
    var x=byRow[q];if(!x){msg.push('baris '+(a+i)+': item tidak dikenal');continue;}
    var nm=x.names.filter(Boolean);
    if(nm.length!==1){msg.push(x.uraian+': item >1 pelaksana / belum ada pelaksana, isi di sheet TARIF_PEKERJAAN');continue;}
    if(!x.id){msg.push(x.uraian+': ID pekerjaan kosong, jalankan Tarif > Perbarui daftar tarif');continue;}
    var val='';
    if(raw!==''&&raw!==null){val=Number(raw);if(typeof raw!=='number'||!isFinite(val)||val<0){msg.push(x.uraian+': harus angka >= 0');continue;}}
    var row=tb.idx[tkey_(x.id,nm[0],x.kode,x.sat)];
    if(row){tsh.getRange(row.r,7).setValue(val);tsh.getRange(row.r,10).setValue(new Date());}
    else add.push([x.id,x.uraian,nm[0],x.kode,x.sat,val,'DETAIL','']);
    ok++;
  }
  tarifTulis_(tsh,add);
  restore();
  if(ok||msg.length)tst((ok?ok+' tarif tersimpan di TARIF_PEKERJAAN (khusus pekerjaan + pelaksana itu). ':'')+(msg.length?'Ditolak: '+msg.slice(0,5).join('; '):''));
}

/* ====================== ENTRI MENU (dipanggil pengguna) ====================== */
function tarifPersiapan(){
  var ui=SpreadsheetApp.getUi();
  tarifKunci_(function(){
    if(tarifAktif_()){ui.alert('Tarif individual sudah AKTIF. Untuk pekerjaan/pelaksana baru pakai "Perbarui daftar tarif".');return;}
    var b=tarifBackup_(),nb=tarifSiapkan_(),d=tarifMigrasi_(true);
    var msg='Backup: '+b+'\nID pekerjaan baru dibuat: '+nb+'\n\nRENCANA MIGRASI (belum ditulis):\n- Baris tarif baru: '+d.baru+'\n- Dari tabel tarif lama: '+d.global+'\n- Dari upah khusus item (V): '+d.v+'\n- Perlu dicek manual: '+d.cek+'\n- Tanpa tarif lama (dibiarkan KOSONG = BELUM DIISI): '+d.kosong+'\n- Sudah ada: '+d.ada+(d.tanpaId.length?'\n- Item tanpa ID (dilewati): '+d.tanpaId.length:'')+'\n\nTulis ke TARIF_PEKERJAAN sekarang?';
    if(ui.alert('A. Persiapan',msg,ui.ButtonSet.YES_NO)!==ui.Button.YES){ui.alert('Dibatalkan. Backup & ID sudah dibuat; tarif belum ditulis.');return;}
    d=tarifMigrasi_(false);
    ui.alert('Migrasi selesai ('+d.baru+' baris). Periksa sheet TARIF_PEKERJAAN (kolom CATATAN), lalu jalankan "B. Validasi".');
  });
}
function tarifValidasi(){
  var ui=SpreadsheetApp.getUi();
  ui.alert('B. Validasi',tarifLaporan_(tarifKunci_(tarifValidasi_)),ui.ButtonSet.OK);
}
function tarifAktifkan(){
  var ui=SpreadsheetApp.getUi();
  if(ui.alert('C. Aktifkan','Pasang rumus tarif individual di NILAI PROGRES (AM:AQ). Validasi dijalankan dulu; bila gagal otomatis dibatalkan. Lanjut?',ui.ButtonSet.YES_NO)!==ui.Button.YES)return;
  ui.alert('C. Aktifkan',tarifKunci_(tarifAktifkan_).msg,ui.ButtonSet.OK);
}
function tarifBangun(){
  var r=tarifKunci_(tarifBangun_);
  SpreadsheetApp.getUi().alert('Daftar tarif diperbarui','Baris baru: '+r.baru+'\nBaris tidak terpakai (ditandai): '+r.orphan,SpreadsheetApp.getUi().ButtonSet.OK);
}
function tarifGantiNama(){
  var ui=SpreadsheetApp.getUi(),a=ui.prompt('Ganti nama pelaksana','Nama LAMA (persis seperti di tarif):',ui.ButtonSet.OK_CANCEL);
  if(a.getSelectedButton()!==ui.Button.OK)return;
  var b=ui.prompt('Ganti nama pelaksana','Nama BARU:',ui.ButtonSet.OK_CANCEL);
  if(b.getSelectedButton()!==ui.Button.OK)return;
  var n=tarifKunci_(function(){return tarifGantiNama_(a.getResponseText(),b.getResponseText());});
  ui.alert(n+' baris tarif dipindah ke nama baru.\nSekarang ganti nama yang sama di BOBOT PEK file STAFF lalu Sinkron. Sebelum itu, tarif tampil BELUM DIISI (tidak salah pindah).');
}
function tarifRollback(){
  var ui=SpreadsheetApp.getUi();
  if(ui.alert('ROLLBACK','Kembalikan rumus tarif lama (AM:AQ dan B50) dari BACKUP_TARIF_ITEM. Sheet TARIF_PEKERJAAN dan tabel tarif lama tidak dihapus. Lanjut?',ui.ButtonSet.YES_NO)!==ui.Button.YES)return;
  tarifKunci_(tarifRollback_);
  ui.alert('Rollback selesai. Tarif kembali memakai tabel lama (pelaksana x satuan).');
}

/* ====================== FUNGSI LAMA YANG DIGANTI ====================== */
/* GANTI onEdit lama dengan ini. Perilaku lama dipertahankan 100% selama tarif individual belum AKTIF. */
function onEdit(e){
  try{
    if(!e||!e.range)return;
    var sh=e.range.getSheet(),nm=sh.getName(),r1=e.range.getRow(),r2=e.range.getLastRow();
    if(nm===TV2.SH){                                   /* catat waktu ubah tarif */
      if(e.range.getColumn()<=7&&e.range.getLastColumn()>=7&&r2>=2){var a=Math.max(2,r1);sh.getRange(a,10,r2-a+1,1).setValue(new Date());}
      return;
    }
    if(nm!=='REKAP PELAKSANA (Rp)')return;
    if(e.range.getColumn()>8||e.range.getLastColumn()<8)return;
    if(tarifAktif_())tarifDariDetailV2_(sh,r1,r2);
    else if(CFG.NSAT)tarifDariDetail_(sh,r1,r2);       /* fungsi lama, tidak diubah */
  }catch(x){SpreadsheetApp.getActive().toast('Sinkron tarif gagal: '+x.message,'Tarif',10);}
}
/* GANTI onOpen lama dengan ini (menu Sinkron tetap, ditambah menu Tarif). */
function onOpen(){
  var ui=SpreadsheetApp.getUi();
  ui.createMenu('Sinkron')
    .addItem('Sinkron sekarang','sinkron')
    .addItem('Pasang sinkron manual (2 arah)','pasangManual')
    .addItem('Manual: tarik dari STAFF','manKeAdmin')
    .addItem('Manual: kirim ke STAFF','manKeStaff')
    .addItem('Aktifkan otomatis tiap 15 menit','pasangOtomatis')
    .addItem('Matikan otomatis','matikanOtomatis')
    .addToUi();
  ui.createMenu('Tarif')
    .addItem('A. Persiapan (backup + ID + migrasi)','tarifPersiapan')
    .addItem('B. Validasi','tarifValidasi')
    .addItem('C. AKTIFKAN tarif individual','tarifAktifkan')
    .addSeparator()
    .addItem('Perbarui daftar tarif','tarifBangun')
    .addItem('Ganti nama pelaksana','tarifGantiNama')
    .addItem('ROLLBACK ke tarif lama','tarifRollback')
    .addToUi();
}
