/* (PEL3 + UPAH V3) SINKRON FILE ADMIN <- FILE STAFF. Upah per pekerjaan+pelaksana di sheet BOBOT PEK (ADMIN), menu Tarif. Tempel di file ADMIN: Extensions > Apps Script (ganti SELURUH isi file lama). */
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

/* ===================== UPAH PER PEKERJAAN + PELAKSANA (V3) =====================
   Sheet BOBOT PEK (ADMIN): 1 baris = 1 pekerjaan; nomor baris SAMA dengan RAB UPDATE / NILAI PROGRES / DATA STAF.
   G:I pelaksana 1 (nama, volume/%, UPAH kuning) ... S:U pelaksana 5. V = CEK. W:AA (tersembunyi) = nama saat upah diisi.
   NILAI PROGRES AM:AQ membaca langsung sel UPAH di baris yang sama (tanpa tabel tarif, tanpa pencarian).
   Upah hanya dipakai bila nama pelaksana sekarang = nama saat upah diisi (mencegah upah salah orang). */
var BP={SH:'BOBOT PEK',AR:'ARSIP_UPGRADE_V3',R0:8,NM:[7,10,13,16,19],VL:[8,11,14,17,20],UP:[9,12,15,18,21],CEK:22,GD:[23,24,25,26,27],RAW:28};

function colL_(n){var s='',m;while(n>0){m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s;}
function bpKunci_(fn){
  var lk=LockService.getDocumentLock();
  if(!lk.tryLock(30000))throw new Error('Spreadsheet sedang dipakai proses lain. Coba lagi sebentar lagi.');
  try{return fn();}finally{lk.releaseLock();}
}
function bpNKS_(){return CFG.NKS!=null?CFG.NKS:3;}
function bpB50_(rp){
  var v=rp.getRange(1,2,Math.min(rp.getLastRow(),150),1).getValues();
  for(var i=0;i<v.length;i++)if(String(v[i][0]).indexOf('Item bertugas tapi tarif belum diisi')===0)return i+1;
  return 0;
}
/* V3 aktif bila rumus AM:AQ membaca sheet BOBOT PEK */
function bpAktif_(){
  var np=SpreadsheetApp.getActive().getSheetByName('NILAI PROGRES');if(!np)return false;
  var F=np.getRange(BP.R0,39,CFG.LASTR-BP.R0+1,1).getFormulas();
  for(var i=0;i<F.length;i++)if(F[i][0]!=='')return String(F[i][0]).indexOf("'BOBOT PEK'!")>=0;
  return false;
}
/* item = baris yang nomor RAB-nya (DATA STAF kolom A) sama dengan nomor barisnya */
function bpItems_(){
  var ss=SpreadsheetApp.getActive(),np=ss.getSheetByName('NILAI PROGRES'),ds=ss.getSheetByName('DATA STAF'),n=CFG.LASTR-BP.R0+1;
  if(!np||!ds)throw new Error('Sheet NILAI PROGRES / DATA STAF tidak ditemukan');
  var V=np.getRange(BP.R0,1,n,48).getValues(),R=ds.getRange(BP.R0,1,n,1).getValues(),out=[],i,k,v,nm;
  var num=function(z){return Number(z)||0;};
  for(i=0;i<n;i++){
    if(Number(R[i][0])!==i+BP.R0)continue;
    v=V[i];nm=[];
    for(k=0;k<5;k++)nm.push(String(v[28+k]==null?'':v[28+k]).trim());
    out.push({r:i+BP.R0,uraian:String(v[1]==null?'':v[1]),st:String(v[18]==null?'':v[18]).trim().toUpperCase(),vol:num(v[2]),names:nm,
      am:[38,39,40,41,42].map(function(z){return num(v[z]);}),ong:[43,44,45,46,47].map(function(z){return num(v[z]);}),oc:num(v[14])});
  }
  return out;
}

/* ---------- rumus sheet BOBOT PEK ---------- */
function bpFA_(r){
  var N="'NILAI PROGRES'!",D="'DATA STAF'!";
  return [
    'IF('+N+'A'+r+'="","",'+N+'A'+r+')',
    'IF('+N+'B'+r+'="","",'+N+'B'+r+')',
    'IF(ISTEXT('+N+'D'+r+'),'+N+'D'+r+',"")',
    'IF(ISNUMBER('+N+'C'+r+'),'+N+'C'+r+',"")',
    'IF('+N+'S'+r+'="","",'+N+'S'+r+')',
    'IF('+D+'$A'+r+'<>'+r+',"",IF(TRIM('+D+'$AK'+r+'&"")="","VOLUME",UPPER(TRIM('+D+'$AK'+r+'&""))))'
  ].map(function(x){return '='+x;});
}
function bpFS_(r,k){
  var N="'NILAI PROGRES'!",D="'DATA STAF'!",a=N+colL_(29+k)+r,nm=colL_(BP.NM[k])+r,v=D+colL_(32+k)+r;
  return ['=IF('+a+'="","",'+a+')','=IF('+nm+'="","",IF('+v+'="","",'+v+'))'];
}
function bpFRaw_(r){
  var nm=[],vl=[],up=[],gd=[],k;
  for(k=0;k<5;k++){nm.push(colL_(BP.NM[k])+r);vl.push(colL_(BP.VL[k])+r);up.push(colL_(BP.UP[k])+r);gd.push(colL_(BP.GD[k])+r);}
  var S='$E'+r,F='$F'+r,D='$D'+r,
    cnt='('+nm.map(function(x){return '('+x+'<>"")';}).join('+')+')',
    a=nm.map(function(x,k){return 'IF(AND('+x+'<>"",'+up[k]+'=""),"UPAH KOSONG "&'+x+'&"; ","")';}),
    b=nm.map(function(x,k){return 'IF(AND('+x+'<>"",'+up[k]+'<>"",TRIM('+gd[k]+'&"")<>TRIM('+x+')),"NAMA BERUBAH "&'+x+'&" (upah dibuat untuk "&IF(TRIM('+gd[k]+'&"")="","?",TRIM('+gd[k]+'&""))&"); ","")';}),
    c=nm.map(function(x,k){return 'IF(AND('+x+'<>"",'+vl[k]+'<>"",ISTEXT('+vl[k]+')),"ANGKA PEMBAGIAN BUKAN ANGKA "&'+x+'&"; ","")';}),
    s='('+nm.map(function(x,k){return 'IF(AND('+x+'<>"",ISNUMBER('+vl[k]+')),'+vl[k]+',0)';}).join('+')+')',
    sp='('+nm.map(function(x,k){return 'IF(AND('+x+'<>"",ISNUMBER('+vl[k]+')),IF('+vl[k]+'>1,'+vl[k]+'/100,'+vl[k]+'),0)';}).join('+')+')',
    p1='IF(AND('+cnt+'=0,'+S+'="AKTIF",N('+D+')>0),"TANPA PELAKSANA; ","")',
    p2='IF(AND('+cnt+'>0,'+S+'="AKTIF",N('+D+')<=0),"VOLUME KOSONG; ","")',
    pm='IF('+cnt+'<2,"",IF('+F+'="PERSEN",IF('+s+'<=0,"PERSEN BELUM DIISI (dibagi rata); ",IF(ABS('+sp+'-1)>0.001,"PERSEN TOTAL "&TEXT('+sp+'*100,"0.#")&"% (harus 100%); ","")),IF('+F+'="SAMA","",IF('+s+'<=0,"VOLUME PELAKSANA BELUM DIISI (dibagi rata); ",IF(ABS('+s+'-N('+D+'))>0.01,"ALOKASI "&TEXT('+s+',"#,##0.##")&" BEDA DARI VOLUME "&TEXT(N('+D+'),"#,##0.##")&"; ","")))))',
    pu='IF(AND('+F+'<>"VOLUME",'+F+'<>"SAMA",'+F+'<>"PERSEN"),"MODE TIDAK DIKENAL; ","")';
  return '=IF(OR('+F+'="",'+S+'="BATAL"),"",'+[p1,p2].concat(a,b,c,[pm,pu]).join('&')+')';
}
function bpFCek_(r){var R='$'+colL_(BP.RAW)+r;return '=IF(OR($F'+r+'="",$E'+r+'="BATAL"),"",IF('+R+'="","OK",LEFT('+R+',LEN('+R+')-2)))';}
function bpFAM_(r,k){
  var nm=colL_(29+k)+r;
  return '=IF(OR('+nm+'="",UPPER(TRIM($S'+r+'))="BATAL"),0,IF(TRIM(\'BOBOT PEK\'!'+colL_(BP.GD[k])+r+'&"")=TRIM('+nm+'),N(\'BOBOT PEK\'!'+colL_(BP.UP[k])+r+'),0))';
}
function bpFB50_(){
  var rg="'BOBOT PEK'!$V$"+BP.R0+':$V$'+CFG.LASTR;
  return '="Item bertugas tapi tarif belum diisi: "&(COUNTIF('+rg+',"*UPAH KOSONG*")+COUNTIF('+rg+',"*NAMA BERUBAH*"))';
}

/* ---------- bangun / perbaiki sheet BOBOT PEK (tidak menyentuh sel UPAH & penanda nama) ---------- */
function bpBangun_(){
  var ss=SpreadsheetApp.getActive(),np=ss.getSheetByName('NILAI PROGRES'),L=CFG.LASTR,R0=BP.R0,n=L-R0+1,sh=ss.getSheetByName(BP.SH),items=bpItems_(),isItem={},k,i,r;
  items.forEach(function(x){isItem[x.r]=1;});
  if(!sh)sh=ss.insertSheet(BP.SH,np.getIndex());
  var fa=[],raw=[],cek=[];
  for(r=R0;r<=L;r++){fa.push(bpFA_(r));raw.push([bpFRaw_(r)]);cek.push([bpFCek_(r)]);}
  sh.getRange(R0,1,n,6).setFormulas(fa);
  for(k=0;k<5;k++){
    var fs=[];for(r=R0;r<=L;r++)fs.push(bpFS_(r,k));
    sh.getRange(R0,BP.NM[k],n,2).setFormulas(fs);
  }
  sh.getRange(R0,BP.CEK,n,1).setFormulas(cek);
  sh.getRange(R0,BP.RAW,n,1).setFormulas(raw);
  /* judul */
  sh.getRange(1,1).setValue('BOBOT PEK + UPAH PELAKSANA (ADMIN)').setFontWeight('bold').setFontSize(13);
  sh.getRange(2,1).setValue('Isi UPAH per satuan (kolom KUNING) untuk tiap pelaksana, di baris pekerjaannya. Nama pelaksana, volume/% dan status otomatis dari file STAFF: jangan diubah di sini. Kolom CEK menjelaskan apa yang perlu dibenahi.');
  sh.getRange(3,1).setValue('Upah dikunci ke NAMA pelaksana saat diisi. Jika nama/urutan pelaksana di STAFF berubah, upah itu TIDAK dipakai dan CEK menandai: periksa, lalu menu Tarif > Konfirmasi upah sesuai nama sekarang.');
  sh.getRange(4,1).setValue('Perlu dicek:').setFontWeight('bold');
  sh.getRange(4,3).setFormula('=SUMPRODUCT((V'+R0+':V'+L+'<>"")*(V'+R0+':V'+L+'<>"OK"))&" item"').setFontWeight('bold').setFontColor('#C00000');
  var h6=['','','','','','','PELAKSANA 1','','','PELAKSANA 2','','','PELAKSANA 3','','','PELAKSANA 4','','','PELAKSANA 5','','','','','','','','',''],
      h7=['No','Uraian pekerjaan','Sat','Volume','Status','Mode'];
  for(k=0;k<5;k++)h7.push('Nama','Volume / %','UPAH /sat (Rp)');
  h7.push('CEK','nama saat upah diisi 1','nama saat upah diisi 2','nama saat upah diisi 3','nama saat upah diisi 4','nama saat upah diisi 5','cek mentah');
  sh.getRange(6,1,1,28).setValues([h6]).setFontWeight('bold').setBackground('#BDD7EE');
  sh.getRange(7,1,1,28).setValues([h7]).setFontWeight('bold').setBackground('#BDD7EE').setWrap(true).setHorizontalAlignment('center');
  /* format */
  var W=[40,300,50,70,70,70];for(i=0;i<6;i++)sh.setColumnWidth(i+1,W[i]);
  for(k=0;k<5;k++){sh.setColumnWidth(BP.NM[k],110);sh.setColumnWidth(BP.VL[k],75);sh.setColumnWidth(BP.UP[k],110);}
  sh.setColumnWidth(BP.CEK,380);
  sh.getRange(R0,4,n,1).setNumberFormat('#,##0.00');
  sh.getRange(R0,2,n,1).setWrap(true);
  var rule=SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).setHelpText('Isi angka rupiah per satuan (>= 0).').build(),unprot=[],rules=[];
  for(k=0;k<5;k++){
    var bg=[];for(r=R0;r<=L;r++)bg.push([isItem[r]?'#FFF2CC':'#FFFFFF']);
    var ur=sh.getRange(R0,BP.UP[k],n,1);
    ur.setBackgrounds(bg).setNumberFormat('#,##0').setDataValidation(rule);
    unprot.push(ur);
    var uc=colL_(BP.UP[k]),nc=colL_(BP.NM[k]),gc=colL_(BP.GD[k]);
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($'+nc+R0+'<>"",$'+uc+R0+'<>"",TRIM($'+gc+R0+'&"")<>TRIM($'+nc+R0+'))').setBackground('#FFD966').setRanges([ur]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($'+nc+R0+'<>"",$'+uc+R0+'="")').setBackground('#FFC7CE').setRanges([ur]).build());
  }
  var cr=sh.getRange(R0,BP.CEK,n,1);
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($V'+R0+'<>"",$V'+R0+'<>"OK")').setBackground('#F8D7DA').setFontColor('#9C0006').setRanges([cr]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$V'+R0+'="OK"').setFontColor('#2E7D32').setRanges([cr]).build());
  sh.setConditionalFormatRules(rules);
  sh.setFrozenRows(7);sh.setFrozenColumns(2);
  sh.hideColumns(BP.GD[0],6);
  /* proteksi lunak: hanya kolom UPAH yang bebas; sel lain memberi peringatan bila diedit */
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p){p.remove();});
  var pr=sh.protect().setDescription('BOBOT PEK: hanya kolom UPAH (kuning) yang diisi').setWarningOnly(true);
  pr.setUnprotectedRanges(unprot);
  /* penghitung di SETUP */
  var su=ss.getSheetByName('SETUP');
  if(su){
    su.getRange(6,1).setValue('BOBOT PEK perlu dicek:').setFontWeight('bold');
    su.getRange(6,2).setFormula("=SUMPRODUCT(('BOBOT PEK'!V"+R0+":V"+L+"<>\"\")*('BOBOT PEK'!V"+R0+":V"+L+"<>\"OK\"))&\" item (lihat sheet BOBOT PEK kolom CEK)\"");
  }
  return sh;
}

/* ---------- arsip kondisi sebelum upgrade (untuk rollback) ---------- */
function bpArsip_(items){
  var ss=SpreadsheetApp.getActive(),np=ss.getSheetByName('NILAI PROGRES'),rp=ss.getSheetByName('REKAP PELAKSANA (Rp)'),n=CFG.LASTR-BP.R0+1,i,k;
  var F=np.getRange(BP.R0,39,n,5).getFormulas(),X=np.getRange(BP.R0,49,n,2).getFormulas(),Vv=np.getRange(BP.R0,39,n,5).getValues();
  var sh=ss.getSheetByName(BP.AR)||ss.insertSheet(BP.AR);sh.clear();
  sh.getRange(1,1,1,13).setValues([['Baris NP','F_AM','F_AN','F_AO','F_AP','F_AQ','F_AW','F_AX','V_AM','V_AN','V_AO','V_AP','V_AQ']]).setFontWeight('bold');
  sh.getRange(2,2,n,7).setNumberFormat('@');                              /* rumus disimpan sebagai TEKS */
  var rows=[];for(i=0;i<n;i++)rows.push([i+BP.R0].concat(F[i],X[i],Vv[i]));
  sh.getRange(2,1,n,13).setValues(rows);
  var b50=bpB50_(rp),fb=b50?rp.getRange(b50,2).getFormula():'',s=0;
  items.forEach(function(x){x.ong.forEach(function(o){s+=o;});});
  sh.getRange(1,15,4,1).setValues([['B50_BARIS'],['B50_RUMUS'],['SUM_ONGKOS'],['WAKTU']]);
  sh.getRange(2,16).setNumberFormat('@');
  sh.getRange(1,16,4,1).setValues([[b50],[fb],[s],[new Date()]]);
  sh.hideSheet();
}

/* ---------- UPGRADE file lama -> BOBOT PEK ---------- */
function bpUpgrade_(){
  var ss=SpreadsheetApp.getActive(),np=ss.getSheetByName('NILAI PROGRES'),rp=ss.getSheetByName('REKAP PELAKSANA (Rp)'),L=CFG.LASTR,n=L-BP.R0+1,k,i,r;
  var already=bpAktif_(),items=bpItems_(),info={impor:0,sudah:already};
  if(!already)bpArsip_(items);
  var sh=bpBangun_();
  if(already)return info;
  /* impor upah dari nilai AM:AQ yang tampil sekarang (hanya sel upah yang masih kosong) */
  var upV=[],gdV=[],byR={},pre=0;
  items.forEach(function(x){byR[x.r]=x;x.ong.forEach(function(o){pre+=o;});});
  for(k=0;k<5;k++){upV.push(sh.getRange(BP.R0,BP.UP[k],n,1).getValues());gdV.push(sh.getRange(BP.R0,BP.GD[k],n,1).getValues());}
  var bad=[];
  items.forEach(function(x){
    for(k=0;k<5;k++){
      var nm=x.names[k];if(!nm||x.st==='BATAL')continue;
      var i0=x.r-BP.R0,cur=upV[k][i0][0],old=x.am[k];
      if(cur===''||cur===null){
        if(old>0){upV[k][i0][0]=old;gdV[k][i0][0]=nm;info.impor++;}
      }else if(Math.abs((Number(cur)||0)-old)>0.005){bad.push('baris '+x.r+' '+nm);}
    }
  });
  if(bad.length)throw new Error('BOBOT PEK sudah berisi upah yang BEDA dari upah lama di: '+bad.slice(0,5).join('; ')+'. Tidak ada yang diubah di NILAI PROGRES. Samakan atau kosongkan sel itu, lalu ulangi.');
  for(k=0;k<5;k++){sh.getRange(BP.R0,BP.UP[k],n,1).setValues(upV[k]);sh.getRange(BP.R0,BP.GD[k],n,1).setValues(gdV[k]);}
  SpreadsheetApp.flush();
  /* pindahkan rumus AM:AQ */
  var F=np.getRange(BP.R0,39,n,5).getFormulas();
  for(i=0;i<n;i++)for(k=0;k<5;k++)if(F[i][k]!=='')F[i][k]=bpFAM_(i+BP.R0,k);
  np.getRange(BP.R0,39,n,5).setFormulas(F);
  np.getRange(BP.R0,49,n,2).clearContent();
  var b50=bpB50_(rp);if(b50)rp.getRange(b50,2).setFormula(bpFB50_());
  bpCatatan_(true);
  SpreadsheetApp.flush();
  /* cek sesudah: upah per pelaksana dan total harus sama persis */
  var after=bpItems_(),post=0,diff=[];
  after.forEach(function(x){x.ong.forEach(function(o){post+=o;});var o=byR[x.r];for(k=0;k<5;k++)if(Math.abs(x.am[k]-o.am[k])>0.005)diff.push('baris '+x.r+' '+x.names[k]);});
  if(diff.length||Math.abs(post-pre)>1){
    bpRollback_();
    throw new Error('Cek sesudah upgrade GAGAL ('+diff.slice(0,5).join('; ')+'; total sebelum '+Math.round(pre)+' sesudah '+Math.round(post)+'). Sudah di-ROLLBACK otomatis.');
  }
  info.total=Math.round(post);
  return info;
}
/* catatan / tanda arsip pada bagian lama (on=true: pasang, false: lepas) */
function bpCatatan_(on){
  var ss=SpreadsheetApp.getActive(),np=ss.getSheetByName('NILAI PROGRES'),rp=ss.getSheetByName('REKAP PELAKSANA (Rp)'),i,nm;
  np.getRange(6,22).setNote(on?'Tidak dipakai lagi. Upah diisi di sheet BOBOT PEK.':'');
  np.getRange(6,23).setNote(on?'Tidak dipakai lagi. Upah diisi di sheet BOBOT PEK.':'');
  var P=null;try{P=lokasiTarif_(rp);}catch(x){}
  if(P){
    var rg=rp.getRange(P.RTH,2,1+CFG.NP,CFG.NSAT+bpNKS_()+1);
    rg.setBackground(on?'#D9D9D9':'#FFFF00');
    rp.getRange(P.RTH,2).setNote(on?'ARSIP: tabel ini tidak dipakai lagi. Upah per pelaksana diisi di sheet BOBOT PEK.':'');
  }
  ['TARIF_PEKERJAAN','BACKUP_TARIF_ITEM','BACKUP_TARIF_GLOBAL'].forEach(function(s){
    var x=ss.getSheetByName(s);if(!x)return;
    try{if(on)x.hideSheet();else x.showSheet();}catch(e){}
  });
}

/* ---------- ROLLBACK ke kondisi sebelum upgrade ---------- */
function bpRollback_(){
  var ss=SpreadsheetApp.getActive(),np=ss.getSheetByName('NILAI PROGRES'),rp=ss.getSheetByName('REKAP PELAKSANA (Rp)'),ar=ss.getSheetByName(BP.AR),n=CFG.LASTR-BP.R0+1;
  if(!ar)throw new Error('Sheet '+BP.AR+' tidak ada: rollback otomatis tidak bisa dilakukan.');
  var d=ar.getRange(2,1,n,8).getValues(),F=[],X=[],i;
  for(i=0;i<n;i++){
    F.push([d[i][1],d[i][2],d[i][3],d[i][4],d[i][5]].map(function(f){return String(f||'');}));
    X.push([String(d[i][6]||''),String(d[i][7]||'')]);
  }
  np.getRange(BP.R0,39,n,5).setFormulas(F);
  np.getRange(BP.R0,49,n,2).setFormulas(X);
  var inf=ar.getRange(1,16,2,1).getValues(),b50=Number(inf[0][0]);
  if(b50&&inf[1][0])rp.getRange(b50,2).setFormula(String(inf[1][0]));
  bpCatatan_(false);
  SpreadsheetApp.flush();
}

/* ---------- KONFIRMASI upah sesuai nama sekarang ---------- */
function bpDaftarBeda_(){
  var ss=SpreadsheetApp.getActive(),sh=ss.getSheetByName(BP.SH),n=CFG.LASTR-BP.R0+1,out=[],k,i;
  if(!sh)return out;
  var nm=[],up=[],gd=[],ur=sh.getRange(BP.R0,2,n,1).getValues();
  for(k=0;k<5;k++){nm.push(sh.getRange(BP.R0,BP.NM[k],n,1).getValues());up.push(sh.getRange(BP.R0,BP.UP[k],n,1).getValues());gd.push(sh.getRange(BP.R0,BP.GD[k],n,1).getValues());}
  for(i=0;i<n;i++)for(k=0;k<5;k++){
    var a=String(nm[k][i][0]).trim(),g=String(gd[k][i][0]).trim();
    if(a!==''&&up[k][i][0]!==''&&a.toLowerCase()!==g.toLowerCase())out.push({r:i+BP.R0,k:k,nama:a,lama:g,uraian:String(ur[i][0]).slice(0,30)});
  }
  return out;
}
function bpKonfirmasi_(){
  var sh=SpreadsheetApp.getActive().getSheetByName(BP.SH),L=bpDaftarBeda_();
  L.forEach(function(x){sh.getRange(x.r,BP.GD[x.k]).setValue(x.nama);});
  return L.length;
}

/* ---------- tanda "upah dibuat untuk <nama>" saat upah diisi ---------- */
function bpStempel_(sh,r1,r2,c1,c2){
  r1=Math.max(r1,BP.R0);if(r2<r1)return;
  var n=r2-r1+1,k,i;
  for(k=0;k<5;k++){
    if(BP.UP[k]<c1||BP.UP[k]>c2)continue;
    var nm=sh.getRange(r1,BP.NM[k],n,1).getValues(),out=[];
    for(i=0;i<n;i++)out.push([String(nm[i][0]).trim()]);
    sh.getRange(r1,BP.GD[k],n,1).setValues(out);
  }
}
/* sheet REKAP PELAKSANA (Rp) saat V3 aktif: kolom Upah satuan (H) hanya tampilan; tabel lama = arsip */
function bpRekapEdit_(sh,r1,r2,c1,c2){
  var ss=SpreadsheetApp.getActive(),tst=function(m){ss.toast(m,'Upah',10);},last=sh.getLastRow(),i;
  if(c1<=8&&c2>=8){
    var Hh=sh.getRange(1,8,last,1).getValues(),d1=-1;
    for(i=0;i<last;i++)if(String(Hh[i][0]).indexOf('Upah satuan')===0){d1=i+2;break;}
    var a=Math.max(r1,d1),b=Math.min(r2,last);
    if(d1>0&&a<=b){
      var n=b-a+1,dF=sh.getRange(a,4,n,1).getFormulas(),cur=sh.getRange(a,8,n,1).getFormulas(),fix=0;
      for(i=0;i<n;i++){
        var mm=/NILAIPROGRES'!F(\d+)$/.exec(String(dF[i][0]).replace(/\s+/g,''));
        if(!mm)continue;
        var f="='NILAI PROGRES'!N"+mm[1];
        if(String(cur[i][0]).replace(/\s+/g,'')!==f){sh.getRange(a+i,8).setFormula(f);fix++;}
      }
      if(fix){tst('Kolom Upah satuan di tabel ini hanya tampilan. Isi upah di sheet BOBOT PEK (kolom kuning).');return;}
    }
  }
  var P=null;try{P=lokasiTarif_(sh);}catch(x){}
  if(P&&r2>=P.RT1&&r1<=P.RT1+CFG.NP-1&&c2>=3&&c1<=2+CFG.NSAT+bpNKS_())tst('Tabel TARIF ini ARSIP dan tidak dipakai lagi. Isi upah di sheet BOBOT PEK.');
}

/* ====================== MENU (dipanggil pengguna) ====================== */
function tarifUpgrade(){
  var ui=SpreadsheetApp.getUi();
  if(bpAktif_()){ui.alert('Sudah memakai sheet BOBOT PEK. Sheet diperiksa/diperbaiki ulang (upah yang sudah diisi tidak diubah).');}
  else if(ui.alert('Upgrade ke BOBOT PEK','Upah pelaksana yang tampil sekarang disalin ke sheet BOBOT PEK, lalu NILAI PROGRES membaca dari sana. Kondisi lama diarsipkan (bisa ROLLBACK). Total upah dicek sebelum dan sesudah; bila beda otomatis dibatalkan. Lanjut?',ui.ButtonSet.YES_NO)!==ui.Button.YES)return;
  var r=bpKunci_(bpUpgrade_);
  ui.alert(r.sudah?'Selesai (sheet diperiksa ulang).':'Upgrade berhasil: '+r.impor+' upah disalin. Total upah individual sama persis ('+r.total+'). Isi/ubah upah di sheet BOBOT PEK; lihat kolom CEK.');
}
function tarifKonfirmasi(){
  var ui=SpreadsheetApp.getUi(),L=bpDaftarBeda_();
  if(!L.length){ui.alert('Semua upah sudah sesuai nama pelaksana sekarang.');return;}
  var t=L.slice(0,10).map(function(x){return 'baris '+x.r+' ('+x.uraian+'): upah dibuat untuk "'+x.lama+'", sekarang "'+x.nama+'"';}).join('\n');
  if(ui.alert('Konfirmasi upah sesuai nama sekarang',L.length+' upah akan dianggap milik nama sekarang:\n'+t+(L.length>10?'\n...':'')+'\n\nPastikan itu memang orang yang sama (bukan tertukar urutan). Lanjut?',ui.ButtonSet.YES_NO)!==ui.Button.YES)return;
  var n=bpKunci_(bpKonfirmasi_);ui.alert(n+' upah dikonfirmasi.');
}
function tarifRollback(){
  var ui=SpreadsheetApp.getUi();
  if(ui.alert('ROLLBACK','Kembalikan rumus upah NILAI PROGRES ke kondisi sebelum upgrade (dari sheet '+BP.AR+'). Sheet BOBOT PEK tidak dihapus. Lanjut?',ui.ButtonSet.YES_NO)!==ui.Button.YES)return;
  bpKunci_(bpRollback_);
  ui.alert('Rollback selesai.');
}

/* ====================== FUNGSI LAMA YANG DIGANTI ====================== */
/* GANTI onEdit lama dengan ini. Perilaku lama (tabel TARIF) tetap berjalan selama BOBOT PEK belum dipakai. */
function onEdit(e){
  try{
    if(!e||!e.range)return;
    var sh=e.range.getSheet(),nm=sh.getName(),r1=e.range.getRow(),r2=e.range.getLastRow(),c1=e.range.getColumn(),c2=e.range.getLastColumn();
    if(nm===BP.SH){bpStempel_(sh,r1,r2,c1,c2);return;}
    if(nm!=='REKAP PELAKSANA (Rp)')return;
    if(bpAktif_()){bpRekapEdit_(sh,r1,r2,c1,c2);return;}
    if(!CFG.NSAT||c1>8||c2<8)return;
    tarifDariDetail_(sh,r1,r2);
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
    .addItem('Konfirmasi upah sesuai nama sekarang','tarifKonfirmasi')
    .addSeparator()
    .addItem('Upgrade file lama ke BOBOT PEK (sekali)','tarifUpgrade')
    .addItem('ROLLBACK upgrade','tarifRollback')
    .addToUi();
}
