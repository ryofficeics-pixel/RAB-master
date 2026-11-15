/* (PEL3 + V4) SINKRON FILE ADMIN <- FILE STAFF. Upah per pekerjaan+pelaksana (BOBOT PEK), material (AHS + MATERIAL), harga online. Tempel di file ADMIN: Extensions > Apps Script (ganti SELURUH isi file lama). */
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

/* ===================== UPAH + MATERIAL + HARGA ONLINE (V4) =====================
   BOBOT PEK (ADMIN): 1 baris = 1 pekerjaan; nomor baris SAMA dengan RAB UPDATE / NILAI PROGRES / DATA STAF.
   G:I pelaksana 1 (nama, volume/%, UPAH kuning) ... S:U pelaksana 5. V = CEK. W:AA (tersembunyi) = nama saat upah diisi.
   AC:AG material kantor/sat, material dibeli pelaksana/sat, RAP item, nilai RAB item, selisih (dari sheet AHS).
   AHS: komponen material per pekerjaan (pilih dari MATERIAL x koefisien). MATERIAL: master harga satuan (+ usulan harga online).
   NILAI PROGRES AM:AQ membaca langsung sel UPAH di baris yang sama. Upah hanya dipakai bila nama pelaksana sekarang = nama saat upah diisi. */
var BP={SH:'BOBOT PEK',AR:'ARSIP_UPGRADE_V3',R0:8,NM:[7,10,13,16,19],VL:[8,11,14,17,20],UP:[9,12,15,18,21],CEK:22,GD:[23,24,25,26,27],RAW:28,MAT:[29,30,31,32,33]};
var MT={SH:'MATERIAL',R0:7,RN:506};
var AH={SH:'AHS',R0:7,RN:1006};

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

function bpAktif_(){
  var np=SpreadsheetApp.getActive().getSheetByName('NILAI PROGRES');if(!np)return false;
  var F=np.getRange(BP.R0,39,CFG.LASTR-BP.R0+1,1).getFormulas();
  for(var i=0;i<F.length;i++)if(F[i][0]!=='')return String(F[i][0]).indexOf("'BOBOT PEK'!")>=0;
  return false;
}

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

function bpFAM_(r,k){
  var nm=colL_(29+k)+r;
  return '=IF(OR('+nm+'="",UPPER(TRIM($S'+r+'))="BATAL"),0,IF(TRIM(\'BOBOT PEK\'!'+colL_(BP.GD[k])+r+'&"")=TRIM('+nm+'),N(\'BOBOT PEK\'!'+colL_(BP.UP[k])+r+'),0))';
}

function bpFB50_(){
  var rg="'BOBOT PEK'!$V$"+BP.R0+':$V$'+CFG.LASTR;
  return '="Item bertugas tapi tarif belum diisi: "&(COUNTIF('+rg+',"*UPAH KOSONG*")+COUNTIF('+rg+',"*NAMA BERUBAH*"))';
}

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

/* ---------- CEK baris BOBOT PEK (kolom mentah AB, ditampilkan di kolom V) ---------- */
function bpFRaw_(r){
  var nm=[],vl=[],up=[],gd=[],k;
  for(k=0;k<5;k++){nm.push(colL_(BP.NM[k])+r);vl.push(colL_(BP.VL[k])+r);up.push(colL_(BP.UP[k])+r);gd.push(colL_(BP.GD[k])+r);}
  var S='$E'+r,F='$F'+r,D='$D'+r,MC=BP.MAT.map(function(c){return colL_(c)+r;}),NN="N('NILAI PROGRES'!N"+r+")",
    cnt='('+nm.map(function(x){return '('+x+'<>"")';}).join('+')+')',
    a=nm.map(function(x,k){return 'IF(AND('+x+'<>"",'+up[k]+'=""),"UPAH KOSONG "&'+x+'&"; ","")';}),
    b=nm.map(function(x,k){return 'IF(AND('+x+'<>"",'+up[k]+'<>"",TRIM('+gd[k]+'&"")<>TRIM('+x+')),"NAMA BERUBAH "&'+x+'&" (upah dibuat untuk "&IF(TRIM('+gd[k]+'&"")="","?",TRIM('+gd[k]+'&""))&"); ","")';}),
    c=nm.map(function(x,k){return 'IF(AND('+x+'<>"",'+vl[k]+'<>"",ISTEXT('+vl[k]+')),"ANGKA PEMBAGIAN BUKAN ANGKA "&'+x+'&"; ","")';}),
    s='('+nm.map(function(x,k){return 'IF(AND('+x+'<>"",ISNUMBER('+vl[k]+')),'+vl[k]+',0)';}).join('+')+')',
    sp='('+nm.map(function(x,k){return 'IF(AND('+x+'<>"",ISNUMBER('+vl[k]+')),IF('+vl[k]+'>1,'+vl[k]+'/100,'+vl[k]+'),0)';}).join('+')+')',
    p1='IF(AND('+cnt+'=0,'+S+'="AKTIF",N('+D+')>0),"TANPA PELAKSANA; ","")',
    p2='IF(AND('+cnt+'>0,'+S+'="AKTIF",N('+D+')<=0),"VOLUME KOSONG; ","")',
    pm='IF('+cnt+'<2,"",IF('+F+'="PERSEN",IF('+s+'<=0,"PERSEN BELUM DIISI (dibagi rata); ",IF(ABS('+sp+'-1)>0.001,"PERSEN TOTAL "&TEXT('+sp+'*100,"0.#")&"% (harus 100%); ","")),IF('+F+'="SAMA","",IF('+s+'<=0,"VOLUME PELAKSANA BELUM DIISI (dibagi rata); ",IF(ABS('+s+'-N('+D+'))>0.01,"ALOKASI "&TEXT('+s+',"#,##0.##")&" BEDA DARI VOLUME "&TEXT(N('+D+'),"#,##0.##")&"; ","")))))',
    pu='IF(AND('+F+'<>"VOLUME",'+F+'<>"SAMA",'+F+'<>"PERSEN"),"MODE TIDAK DIKENAL; ","")',
    p3='IF(AND(N('+MC[1]+')>0,'+NN+'>0,'+NN+'<N('+MC[1]+')),"UPAH/BORONGAN LEBIH KECIL DARI MATERIAL PELAKSANA; ","")',
    p4='IF(AND(N('+MC[3]+')>0,N('+MC[4]+')<0),"RAP LEBIH BESAR DARI NILAI RAB; ","")';
  return '=IF(OR('+F+'="",'+S+'="BATAL"),"",'+[p1,p2].concat(a,b,c,[pm,pu,p3,p4]).join('&')+')';
}
function bpFCek_(r){var R='$'+colL_(BP.RAW)+r;return '=IF(OR($F'+r+'="",$E'+r+'="BATAL"),"",IF('+R+'="","OK",LEFT('+R+',LEN('+R+')-2)))';}
/* material & RAP per pekerjaan: AC material kantor/sat, AD material dibeli pelaksana/sat, AE RAP item, AF nilai RAB item, AG selisih */
function bpFMat_(r){
  var H='AHS!$H$'+AH.R0+':$H$'+AH.RN,A='AHS!$A$'+AH.R0+':$A$'+AH.RN,G='AHS!$G$'+AH.R0+':$G$'+AH.RN,
    C=BP.MAT.map(function(c){return colL_(c)+r;}),F='$F'+r,live='OR('+F+'="",$E'+r+'="BATAL")';
  return [
    '=IF('+F+'="","",SUMIFS('+H+','+A+','+r+')-'+C[1]+')',
    '=IF('+F+'="","",SUMIFS('+H+','+A+','+r+','+G+',"PELAKSANA"))',
    '=IF('+live+',"",(N(\'NILAI PROGRES\'!N'+r+')+'+C[0]+')*N($D'+r+'))',
    '=IF('+live+',"",N(\'NILAI PROGRES\'!F'+r+'))',
    '=IF('+live+',"",'+C[3]+'-'+C[2]+')'
  ];
}

/* ---------- bangun / perbaiki BOBOT PEK + AHS + MATERIAL (tidak menyentuh sel input) ---------- */
function bpSheet_(nm,after){var ss=SpreadsheetApp.getActive();return ss.getSheetByName(nm)||ss.insertSheet(nm,after.getIndex());}
function bpBangun_(){
  var ss=SpreadsheetApp.getActive(),np=ss.getSheetByName('NILAI PROGRES'),L=CFG.LASTR,R0=BP.R0,n=L-R0+1,items=bpItems_(),isItem={},k,i,r;
  items.forEach(function(x){isItem[x.r]=1;});
  var sh=bpSheet_(BP.SH,np),ah=bpSheet_(AH.SH,sh),mt=bpSheet_(MT.SH,ah);        /* semua sheet ada dulu sebelum rumus saling merujuk */
  var fa=[],raw=[],cek=[],mf=[];
  for(r=R0;r<=L;r++){fa.push(bpFA_(r));raw.push([bpFRaw_(r)]);cek.push([bpFCek_(r)]);mf.push(bpFMat_(r));}
  sh.getRange(R0,1,n,6).setFormulas(fa);
  for(k=0;k<5;k++){
    var fs=[];for(r=R0;r<=L;r++)fs.push(bpFS_(r,k));
    sh.getRange(R0,BP.NM[k],n,2).setFormulas(fs);
  }
  sh.getRange(R0,BP.CEK,n,1).setFormulas(cek);
  sh.getRange(R0,BP.RAW,n,1).setFormulas(raw);
  sh.getRange(R0,BP.MAT[0],n,5).setFormulas(mf);
  sh.getRange(1,1).setValue('BOBOT PEK + UPAH PELAKSANA + RAP (ADMIN)').setFontWeight('bold').setFontSize(13);
  sh.getRange(2,1).setValue('Isi UPAH per satuan (kolom KUNING) untuk tiap pelaksana, di baris pekerjaannya. Nama, volume/% dan status otomatis dari file STAFF: jangan diubah di sini. MATERIAL: isi komponen di sheet AHS (harga dari sheet MATERIAL); hasilnya tampil di kolom hijau. Kolom CEK menjelaskan yang perlu dibenahi.');
  sh.getRange(3,1).setValue('Upah dikunci ke NAMA pelaksana saat diisi. Jika nama/urutan pelaksana di STAFF berubah, upah itu TIDAK dipakai dan CEK menandai: periksa, lalu menu Tarif > Konfirmasi upah sesuai nama sekarang.');
  sh.getRange(4,1).setValue('Perlu dicek:').setFontWeight('bold');
  sh.getRange(4,3).setFormula('=SUMPRODUCT((V'+R0+':V'+L+'<>"")*(V'+R0+':V'+L+'<>"OK"))&" item"').setFontWeight('bold').setFontColor('#C00000');
  var h6=['','','','','','','PELAKSANA 1','','','PELAKSANA 2','','','PELAKSANA 3','','','PELAKSANA 4','','','PELAKSANA 5','','','','','','','','',''],
      h7=['No','Uraian pekerjaan','Sat','Volume','Status','Mode'];
  for(k=0;k<5;k++)h7.push('Nama','Volume / %','UPAH /sat (Rp)');
  h7.push('CEK','nama saat upah diisi 1','nama saat upah diisi 2','nama saat upah diisi 3','nama saat upah diisi 4','nama saat upah diisi 5','cek mentah');
  sh.getRange(6,1,1,28).setValues([h6]).setFontWeight('bold').setBackground('#BDD7EE');
  sh.getRange(7,1,1,28).setValues([h7]).setFontWeight('bold').setBackground('#BDD7EE').setWrap(true).setHorizontalAlignment('center');
  var mc=BP.MAT[0],mL=colL_;
  sh.getRange(6,mc,1,5).setValues([['MATERIAL & RAP (material dari sheet AHS)','','','','']]).setFontWeight('bold').setBackground('#C6E0B4');
  sh.getRange(7,mc,1,5).setValues([['MATERIAL KANTOR /sat (Rp)','MATERIAL DIBELI PELAKSANA /sat (Rp)','RAP item (Rp) = (upah + material kantor) x volume','NILAI RAB item (Rp)','SELISIH RAB - RAP (Rp)']]).setFontWeight('bold').setBackground('#C6E0B4').setWrap(true).setHorizontalAlignment('center');
  sh.getRange(5,mc+1).setValue('TOTAL (tanpa item BATAL)').setFontWeight('bold');
  sh.getRange(5,mc+2,1,3).setFormulas([['=SUM('+mL(mc+2)+R0+':'+mL(mc+2)+L+')','=SUM('+mL(mc+3)+R0+':'+mL(mc+3)+L+')','=SUM('+mL(mc+4)+R0+':'+mL(mc+4)+L+')']]).setFontWeight('bold').setNumberFormat('#,##0');
  var W=[40,300,50,70,70,70];for(i=0;i<6;i++)sh.setColumnWidth(i+1,W[i]);
  for(k=0;k<5;k++){sh.setColumnWidth(BP.NM[k],110);sh.setColumnWidth(BP.VL[k],75);sh.setColumnWidth(BP.UP[k],110);}
  sh.setColumnWidth(BP.CEK,380);
  for(i=0;i<5;i++)sh.setColumnWidth(mc+i,125);
  sh.getRange(R0,4,n,1).setNumberFormat('#,##0.00');
  sh.getRange(R0,mc,n,5).setNumberFormat('#,##0');
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
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER($'+mL(mc+4)+R0+'),$'+mL(mc+4)+R0+'<0)').setFontColor('#9C0006').setRanges([sh.getRange(R0,mc+4,n,1)]).build());
  sh.setConditionalFormatRules(rules);
  sh.setFrozenRows(7);sh.setFrozenColumns(2);
  sh.hideColumns(BP.GD[0],6);
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p){p.remove();});
  var pr=sh.protect().setDescription('BOBOT PEK: hanya kolom UPAH (kuning) yang diisi').setWarningOnly(true);
  pr.setUnprotectedRanges(unprot);
  var su=ss.getSheetByName('SETUP');
  if(su){
    su.getRange(6,1).setValue('BOBOT PEK perlu dicek:').setFontWeight('bold');
    su.getRange(6,2).setFormula("=SUMPRODUCT(('BOBOT PEK'!V"+R0+":V"+L+"<>\"\")*('BOBOT PEK'!V"+R0+":V"+L+"<>\"OK\"))&\" item (lihat sheet BOBOT PEK kolom CEK)\"");
  }
  matBangun_(mt);ahsBangun_(ah);
  return sh;
}

/* ---------- sheet MATERIAL: master harga satuan material (satu-satunya tempat harga diubah) ---------- */
function matF_(r){
  var R0=MT.R0,RN=MT.RN;
  return {
    G:'=IF(OR($F'+r+'="",N($C'+r+')=0),"",$F'+r+'/$C'+r+'-1)',
    J:'=IF($A'+r+'="","",IF(COUNTIF($A$'+R0+':$A$'+RN+',$A'+r+')>1,"NAMA GANDA",IF($B'+r+'="","SATUAN KOSONG",IF(NOT(ISNUMBER($C'+r+')),"HARGA KOSONG","OK"))))'
  };
}
function matBangun_(sh){
  var R0=MT.R0,RN=MT.RN,n=RN-R0+1,r,G=[],J=[];
  for(r=R0;r<=RN;r++){var f=matF_(r);G.push([f.G]);J.push([f.J]);}
  sh.getRange(R0,7,n,1).setFormulas(G);
  sh.getRange(R0,10,n,1).setFormulas(J);
  sh.getRange(1,1).setValue('MATERIAL: DAFTAR HARGA SATUAN (master)').setFontWeight('bold').setFontSize(13);
  sh.getRange(2,1).setValue('Satu baris = satu material dengan satuan belinya (sak, kg, m3, meter, buah, ...). Ubah HARGA di sini saja: semua AHS dan RAP ikut berubah. Tambah material baru di baris kosong. Sheet AHS memilih material dari daftar ini.');
  sh.getRange(3,1).setValue('Wilayah pencarian harga online:').setFontWeight('bold');
  if(String(sh.getRange(3,3).getValue()).trim()==='')sh.getRange(3,3).setValue('Indonesia');
  sh.getRange(3,3).setBackground('#FFF2CC');
  sh.getRange(3,4).setValue('Isi YA di kolom E untuk material yang harganya ingin dicari online (menu Material > Sync harga online). Hasil masuk ke kolom USULAN dan baru dipakai setelah "Terapkan usulan".');
  sh.getRange(4,1).setValue('Perlu dicek:').setFontWeight('bold');
  sh.getRange(4,3).setFormula('=SUMPRODUCT((J'+R0+':J'+RN+'<>"")*(J'+R0+':J'+RN+'<>"OK"))&" baris"').setFontWeight('bold').setFontColor('#C00000');
  sh.getRange(4,6).setFormula('=(COUNTIF(G'+R0+':G'+RN+',">=0.1")+COUNTIF(G'+R0+':G'+RN+',"<=-0.1"))&" usulan beda >= 10%"').setFontWeight('bold');
  sh.getRange(6,1,1,10).setValues([['Material','Sat','HARGA (Rp/sat)','Update harga','CARI ONLINE? (YA)','Harga usulan online','Selisih usulan','Tgl cek online','Sumber / catatan online','CEK']]).setFontWeight('bold').setBackground('#BDD7EE').setWrap(true).setHorizontalAlignment('center');
  var W=[230,70,120,100,90,120,90,100,420,150];for(var i=0;i<10;i++)sh.setColumnWidth(i+1,W[i]);
  [1,2,3,5].forEach(function(c){sh.getRange(R0,c,n,1).setBackground('#FFF2CC');});
  sh.getRange(R0,3,n,1).setNumberFormat('#,##0').setDataValidation(SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).setHelpText('Isi angka rupiah per satuan.').build());
  sh.getRange(R0,4,n,1).setNumberFormat('dd-mmm-yyyy');
  sh.getRange(R0,6,n,1).setNumberFormat('#,##0');
  sh.getRange(R0,7,n,1).setNumberFormat('+0.0%;-0.0%');
  sh.getRange(R0,8,n,1).setNumberFormat('dd-mmm-yyyy');
  sh.getRange(R0,5,n,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['YA'],true).setAllowInvalid(true).build());
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($J'+R0+'<>"",$J'+R0+'<>"OK")').setBackground('#F8D7DA').setFontColor('#9C0006').setRanges([sh.getRange(R0,10,n,1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER($G'+R0+'),ABS($G'+R0+')>=0.1)').setBackground('#FFD966').setRanges([sh.getRange(R0,6,n,2)]).build()]);
  sh.setFrozenRows(6);sh.setFrozenColumns(1);
  return sh;
}

/* ---------- sheet AHS: komponen material per pekerjaan (material dipilih dari MATERIAL x koefisien) ---------- */
function ahsF_(r){
  var BN="'BOBOT PEK'!",MA='MATERIAL!$A$'+MT.R0+':$A$'+MT.RN,MB='MATERIAL!$B$'+MT.R0+':$B$'+MT.RN,MC='MATERIAL!$C$'+MT.R0+':$C$'+MT.RN,
    vol='N(IFERROR(INDEX('+BN+'$D:$D,$A'+r+'),0))';
  return {
    B:'=IF($A'+r+'="","",IF(N($A'+r+')<8,"?",IFERROR(INDEX('+BN+'$B:$B,$A'+r+'),"?")))',
    D:'=IF($C'+r+'="","",IFERROR(INDEX('+MB+',MATCH($C'+r+','+MA+',0)),"?"))',
    F:'=IF($C'+r+'="","",IFERROR(N(INDEX('+MC+',MATCH($C'+r+','+MA+',0))),"?"))',
    H:'=IF($A'+r+'="","",IF(AND(ISNUMBER($E'+r+'),N($F'+r+')>0),$E'+r+'*$F'+r+',0))',
    I:'=IF(OR($A'+r+'="",NOT(ISNUMBER($E'+r+'))),"",$E'+r+'*'+vol+')',
    J:'=IF(OR($A'+r+'="",$H'+r+'=""),"",$H'+r+'*'+vol+')',
    K:'=IF(AND($A'+r+'="",$C'+r+'="",$E'+r+'=""),"",IF(N($A'+r+')<8,"ISI NO BARIS (angka baris di sheet BOBOT PEK)",IF(IFERROR(INDEX('+BN+'$F:$F,$A'+r+'),"")="","BARIS BUKAN PEKERJAAN",IF($C'+r+'="","MATERIAL KOSONG",IF($D'+r+'="?","MATERIAL TIDAK ADA DI SHEET MATERIAL",IF(NOT(ISNUMBER($E'+r+')),"KOEFISIEN KOSONG",IF(N($F'+r+')<=0,"HARGA MATERIAL KOSONG (isi di sheet MATERIAL)",IF(AND($G'+r+'<>"",$G'+r+'<>"KANTOR",$G'+r+'<>"PELAKSANA"),"DIBELI OLEH: isi KANTOR atau PELAKSANA","OK"))))))))'
  };
}
function ahsBangun_(sh){
  var R0=AH.R0,RN=AH.RN,n=RN-R0+1,r,B=[],D=[],F=[],H=[],I=[],J=[],K=[];
  for(r=R0;r<=RN;r++){var f=ahsF_(r);B.push([f.B]);D.push([f.D]);F.push([f.F]);H.push([f.H]);I.push([f.I]);J.push([f.J]);K.push([f.K]);}
  sh.getRange(R0,2,n,1).setFormulas(B);
  sh.getRange(R0,4,n,1).setFormulas(D);
  sh.getRange(R0,6,n,1).setFormulas(F);
  sh.getRange(R0,8,n,1).setFormulas(H);
  sh.getRange(R0,9,n,1).setFormulas(I);
  sh.getRange(R0,10,n,1).setFormulas(J);
  sh.getRange(R0,11,n,1).setFormulas(K);
  sh.getRange(1,1).setValue('AHS: MATERIAL PER PEKERJAAN (material x koefisien)').setFontWeight('bold').setFontSize(13);
  sh.getRange(2,1).setValue('Satu baris = satu komponen material pada satu pekerjaan. NO BARIS = angka baris pekerjaan di sheet BOBOT PEK (uraiannya tampil di kolom B untuk dicek). MATERIAL = pilih dari sheet MATERIAL (satuan & harga otomatis). KOEFISIEN = jumlah material untuk 1 SATUAN PEKERJAAN.');
  sh.getRange(3,1).setValue('DIBELI OLEH: KANTOR (masuk RAP kantor) atau PELAKSANA (borongan: material dibeli pelaksana/vendor; hanya pembanding terhadap upah). Kosong = KANTOR. Material tidak dibagi per pelaksana. Contoh: pasang titik lampu = 3 baris (kabel meter, fitting unit, lampu unit) dengan NO BARIS yang sama.');
  sh.getRange(4,1).setValue('Baris perlu dicek:').setFontWeight('bold');
  sh.getRange(4,3).setFormula('=SUMPRODUCT((K'+R0+':K'+RN+'<>"")*(K'+R0+':K'+RN+'<>"OK"))&" baris"').setFontWeight('bold').setFontColor('#C00000');
  sh.getRange(6,1,1,11).setValues([['NO BARIS (BOBOT PEK)','Uraian pekerjaan','Material','Sat material','KOEFISIEN per 1 sat pekerjaan','Harga /sat material (Rp)','DIBELI OLEH','Biaya per 1 sat pekerjaan (Rp)','Kebutuhan total (qty material)','Biaya total (Rp)','CEK']]).setFontWeight('bold').setBackground('#BDD7EE').setWrap(true).setHorizontalAlignment('center');
  var W=[80,300,200,80,100,120,100,120,120,120,320];for(var i=0;i<11;i++)sh.setColumnWidth(i+1,W[i]);
  [1,3,5,7].forEach(function(c){sh.getRange(R0,c,n,1).setBackground('#FFF2CC');});
  sh.getRange(R0,5,n,1).setNumberFormat('#,##0.####');
  [6,8,10].forEach(function(c){sh.getRange(R0,c,n,1).setNumberFormat('#,##0');});
  sh.getRange(R0,9,n,1).setNumberFormat('#,##0.##');
  var ss=SpreadsheetApp.getActive();
  sh.getRange(R0,1,n,1).setDataValidation(SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(8).setAllowInvalid(false).setHelpText('Isi angka baris pekerjaan di sheet BOBOT PEK.').build());
  sh.getRange(R0,3,n,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(ss.getSheetByName(MT.SH).getRange(MT.R0,1,MT.RN-MT.R0+1,1),true).setAllowInvalid(false).setHelpText('Pilih dari sheet MATERIAL (tambahkan dulu di sana bila belum ada).').build());
  sh.getRange(R0,5,n,1).setDataValidation(SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).setHelpText('Isi angka >= 0.').build());
  sh.getRange(R0,7,n,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['KANTOR','PELAKSANA'],true).setAllowInvalid(false).build());
  sh.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($K'+R0+'<>"",$K'+R0+'<>"OK")').setBackground('#F8D7DA').setFontColor('#9C0006').setRanges([sh.getRange(R0,11,n,1)]).build()]);
  sh.setFrozenRows(6);
  return sh;
}

/* ---------- bagian lama yang disamarkan setelah BOBOT PEK aktif ---------- */
function bpTabelLama_(rp){
  var B=rp.getRange(1,2,Math.max(rp.getLastRow(),1),1).getValues(),i;
  for(i=0;i<B.length;i++)if(String(B[i][0]).indexOf('TARIF UPAH PER PELAKSANA')===0)return {RTB:i+1,RTH:i+2,RT1:i+3};
  return null;
}
function bpCatatan_(on){
  var ss=SpreadsheetApp.getActive(),np=ss.getSheetByName('NILAI PROGRES'),rp=ss.getSheetByName('REKAP PELAKSANA (Rp)'),T=bpTabelLama_(rp);
  np.getRange(6,22).setNote(on?'Tidak dipakai lagi. Upah diisi di sheet BOBOT PEK.':'');
  np.getRange(6,23).setNote(on?'Tidak dipakai lagi. Upah diisi di sheet BOBOT PEK.':'');
  if(T){
    rp.getRange(T.RTH,2,1+CFG.NP,CFG.NSAT+bpNKS_()+1).setBackground(on?'#D9D9D9':'#FFFF00');
    rp.getRange(T.RTH,2).setNote(on?'ARSIP: tabel ini tidak dipakai lagi. Upah per pelaksana diisi di sheet BOBOT PEK.':'');
  }
  ['TARIF_PEKERJAAN','BACKUP_TARIF_ITEM','BACKUP_TARIF_GLOBAL'].forEach(function(s){
    var x=ss.getSheetByName(s);if(!x)return;
    try{if(on)x.hideSheet();else x.showSheet();}catch(e){}
  });
}
/* HAPUS PERMANEN sisa versi lama (setelah V3/V4 dipastikan benar). Rollback tidak bisa lagi sesudahnya. */
function bpBersih_(){
  var ss=SpreadsheetApp.getActive(),out=[],np=ss.getSheetByName('NILAI PROGRES'),ru=ss.getSheetByName('RAB UPDATE'),rp=ss.getSheetByName('REKAP PELAKSANA (Rp)');
  if(!bpAktif_())throw new Error('BOBOT PEK belum aktif. Jalankan "Perbarui BOBOT PEK + AHS + MATERIAL" dulu.');
  ['TARIF_PEKERJAAN','BACKUP_TARIF_ITEM','BACKUP_TARIF_GLOBAL',BP.AR].forEach(function(s){var x=ss.getSheetByName(s);if(x){ss.deleteSheet(x);out.push('sheet '+s);}});
  if(ru){
    var qv=ru.getRange(1,17,Math.max(ru.getLastRow(),1),1).getValues(),ch=0;
    qv.forEach(function(r){if(/^PK-[0-9A-F]{8}$/i.test(String(r[0]))||String(r[0]).indexOf('ID PEKERJAAN')===0){r[0]='';ch++;}});
    if(ch){ru.getRange(1,17,qv.length,1).setValues(qv);out.push('RAB UPDATE kolom Q ('+ch+' ID lama)');}
  }
  if(np&&String(np.getRange(6,49).getValue()).indexOf('ID PEKERJAAN')===0){np.getRange(6,49,1,2).clearContent();out.push('NILAI PROGRES AW:AX (judul lama)');}
  var T=rp?bpTabelLama_(rp):null;
  if(T){var nr=CFG.NP+2;rp.getRange(T.RTB,2,nr,CFG.NSAT+bpNKS_()+1).clear();rp.hideRows(T.RTB,nr);out.push('tabel TARIF lama di REKAP (dikosongkan & disembunyikan)');}
  return out;
}
/* sheet REKAP saat BOBOT PEK aktif: kolom Upah satuan (H) hanya tampilan; tabel lama = arsip */
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
  var T=bpTabelLama_(sh);
  if(T&&r2>=T.RT1&&r1<=T.RT1+CFG.NP-1&&c2>=3&&c1<=2+CFG.NSAT+bpNKS_())tst('Tabel TARIF ini ARSIP dan tidak dipakai lagi. Isi upah di sheet BOBOT PEK.');
}
/* tanggal "Update harga" otomatis saat HARGA di sheet MATERIAL diubah */
function matStempel_(sh,r1,r2,c1,c2){
  r1=Math.max(r1,MT.R0);if(r2<r1||c1>3||c2<3)return;
  var n=r2-r1+1,v=sh.getRange(r1,3,n,1).getValues(),out=[],i,now=new Date(),d=sh.getRange(r1,4,n,1).getValues();
  for(i=0;i<n;i++)out.push([v[i][0]===''?'':now]);
  sh.getRange(r1,4,n,1).setValues(out);
}

/* ====================== HARGA ONLINE (usulan harga, tidak menimpa harga master) ====================== */
function hargaNum_(raw){
  var m=/^(\d{1,3}(?:[.,]\d{3})*)[.,](\d{1,2})$/.exec(raw);
  if(m)return Number(m[1].replace(/[.,]/g,'')+'.'+m[2]);
  if(/^\d{1,3}([.,]\d{3})+$/.test(raw)||/^\d+$/.test(raw))return Number(raw.replace(/[.,]/g,''));
  return NaN;
}
function hargaParseRp_(t){
  var re=/Rp\.?\s*([0-9]+(?:[.,][0-9]+)*)\s*(juta|jt|ribu|rb|k)?/gi,m,out=[],v,u;
  t=String(t||'');
  while((m=re.exec(t))){
    v=hargaNum_(m[1]);u=(m[2]||'').toLowerCase();
    if(u==='juta'||u==='jt')v*=1e6;else if(u==='ribu'||u==='rb'||u==='k')v*=1e3;
    if(v>0&&isFinite(v))out.push(v);
  }
  return out;
}
function hargaMedian_(a){a=a.slice().sort(function(x,y){return x-y;});var n=a.length;return n?(n%2?a[(n-1)/2]:(a[n/2-1]+a[n/2])/2):0;}
function hargaRingkas_(v){                                  /* median tanpa pencilan (di luar 0,4x - 2,5x median) */
  if(!v.length)return null;
  var m=hargaMedian_(v),f=v.filter(function(x){return x>=m*0.4&&x<=m*2.5;});
  return f.length?{harga:Math.round(hargaMedian_(f)),n:f.length}:null;
}
function hargaJson_(t){
  var a=String(t||'').indexOf('{'),b=String(t||'').lastIndexOf('}'),o;
  if(a<0||b<=a)return {harga:null,n:0,sumber:[],catatan:'balasan bukan JSON'};
  try{o=JSON.parse(String(t).slice(a,b+1));}catch(e){return {harga:null,n:0,sumber:[],catatan:'JSON tidak valid'};}
  var h=Number(o.harga);
  return {harga:(isFinite(h)&&h>0&&h<1e9)?Math.round(h):null,n:Number(o.n)||0,sumber:(o.sumber||[]).slice(0,3),catatan:String(o.catatan||'')};
}
function hargaCariClaude_(nama,sat,wil,P){
  var q='Cari harga eceran terbaru material konstruksi "'+nama+'" per '+sat+' di '+wil+' (Indonesia) dengan pencarian web. Balas HANYA satu objek JSON: {"harga":<angka rupiah per '+sat+', atau null>,"n":<jumlah sumber yang dipakai>,"sumber":[<url>],"catatan":"<singkat; sebut bila satuan sumber dikonversi>"}. Jangan menebak: bila tidak ada harga yang jelas, harga null.';
  var res=UrlFetchApp.fetch('https://api.anthropic.com/v1/messages',{method:'post',contentType:'application/json',muteHttpExceptions:true,
    headers:{'x-api-key':P.key,'anthropic-version':'2023-06-01'},
    payload:JSON.stringify({model:P.model,max_tokens:1024,tools:[{type:'web_search_20250305',name:'web_search',max_uses:3}],messages:[{role:'user',content:q}]})});
  if(res.getResponseCode()!==200){var e=new Error('Claude API '+res.getResponseCode()+': '+String(res.getContentText()).slice(0,160));e.fatal=(res.getResponseCode()===401||res.getResponseCode()===403||res.getResponseCode()===429);throw e;}
  var j=JSON.parse(res.getContentText()),t=(j.content||[]).filter(function(b){return b.type==='text';}).map(function(b){return b.text;}).join('\n');
  return hargaJson_(t);
}
function hargaCariGoogle_(nama,sat,wil,P){
  var u='https://www.googleapis.com/customsearch/v1?key='+encodeURIComponent(P.key)+'&cx='+encodeURIComponent(P.cx)+'&num=10&q='+encodeURIComponent(nama+' '+sat+' harga '+wil);
  var res=UrlFetchApp.fetch(u,{muteHttpExceptions:true});
  if(res.getResponseCode()!==200){var e=new Error('Google CSE '+res.getResponseCode()+': '+String(res.getContentText()).slice(0,160));e.fatal=(res.getResponseCode()===400||res.getResponseCode()===403||res.getResponseCode()===429);throw e;}
  var it=(JSON.parse(res.getContentText()).items||[]),vals=[],src=[];
  it.forEach(function(x){var p=hargaParseRp_((x.title||'')+' '+(x.snippet||''));if(p.length){vals.push(p[0]);src.push(x.link);}});
  var r=hargaRingkas_(vals);
  return r&&r.n>=2?{harga:r.harga,n:r.n,sumber:src.slice(0,3),catatan:'dari cuplikan hasil pencarian; CEK SATUAN'}:{harga:null,n:vals.length,sumber:[],catatan:'harga tidak cukup jelas di hasil pencarian'};
}
function hargaProp_(){
  var P=PropertiesService.getScriptProperties();
  return {prov:P.getProperty('HARGA_PROVIDER'),key:P.getProperty('HARGA_API_KEY'),cx:P.getProperty('HARGA_CSE_CX'),model:P.getProperty('HARGA_MODEL')||'claude-haiku-5-5'};
}
function hargaSync_(maks){
  var sh=SpreadsheetApp.getActive().getSheetByName(MT.SH),P=hargaProp_(),R0=MT.R0,n=MT.RN-R0+1;
  if(!sh)throw new Error('Sheet MATERIAL belum ada. Jalankan "Perbarui BOBOT PEK + AHS + MATERIAL".');
  if(!P.prov||!P.key||(P.prov==='GOOGLE'&&!P.cx))throw new Error('Koneksi harga belum diatur: menu Material > Atur koneksi harga online.');
  var v=sh.getRange(R0,1,n,9).getValues(),wil=String(sh.getRange(3,3).getValue()||'Indonesia'),T=[],i;
  for(i=0;i<n;i++)if(String(v[i][0]).trim()!==''&&String(v[i][1]).trim()!==''&&String(v[i][4]).trim().toUpperCase()==='YA')T.push({r:R0+i,nama:String(v[i][0]).trim(),sat:String(v[i][1]).trim(),t:v[i][7]instanceof Date?v[i][7].getTime():0});
  T.sort(function(a,b){return a.t-b.t;});
  var run=T.slice(0,maks||10),out={dicek:0,ketemu:0,gagal:0,sisa:T.length-run.length,galat:''},now=new Date();
  for(i=0;i<run.length;i++){
    var x=run[i],res;
    try{res=P.prov==='GOOGLE'?hargaCariGoogle_(x.nama,x.sat,wil,P):hargaCariClaude_(x.nama,x.sat,wil,P);}
    catch(e){out.gagal++;sh.getRange(x.r,9).setValue('ERROR: '+e.message);if(e.fatal){out.galat=e.message;out.sisa+=run.length-i-1;break;}continue;}
    out.dicek++;
    sh.getRange(x.r,8).setValue(now);
    if(res.harga){out.ketemu++;sh.getRange(x.r,6).setValue(res.harga);sh.getRange(x.r,9).setValue(res.n+' sumber: '+res.sumber.join(' | ')+(res.catatan?' | '+res.catatan:''));}
    else{sh.getRange(x.r,6).setValue('');sh.getRange(x.r,9).setValue('TIDAK DITEMUKAN: '+res.catatan);}
  }
  return out;
}
function hargaDaftarUsulan_(){
  var sh=SpreadsheetApp.getActive().getSheetByName(MT.SH),n=MT.RN-MT.R0+1,out=[],i;
  if(!sh)return out;
  var v=sh.getRange(MT.R0,1,n,6).getValues();
  for(i=0;i<n;i++){
    var c=Number(v[i][2])||0,f=v[i][5];
    if(String(v[i][0]).trim()!==''&&typeof f==='number'&&f>0&&Math.abs(f-c)>0.5)out.push({r:MT.R0+i,nama:String(v[i][0]),sat:String(v[i][1]),lama:c,baru:f});
  }
  return out;
}
function hargaTerapkan_(){
  var sh=SpreadsheetApp.getActive().getSheetByName(MT.SH),L=hargaDaftarUsulan_(),now=new Date();
  L.forEach(function(x){sh.getRange(x.r,3).setValue(x.baru);sh.getRange(x.r,4).setValue(now);sh.getRange(x.r,6).setValue('');});
  return L.length;
}

/* ====================== MENU (dipanggil pengguna) ====================== */
function tarifUpgrade(){
  var ui=SpreadsheetApp.getUi();
  if(bpAktif_()){ui.alert('Sudah memakai sheet BOBOT PEK. BOBOT PEK, AHS dan MATERIAL diperiksa/dilengkapi ulang (upah, baris AHS dan harga material yang sudah diisi tidak diubah).');}
  else if(ui.alert('Upgrade ke BOBOT PEK','Upah pelaksana yang tampil sekarang disalin ke sheet BOBOT PEK, lalu NILAI PROGRES membaca dari sana. Sheet AHS dan MATERIAL ikut dibuat. Kondisi lama diarsipkan (bisa ROLLBACK). Total upah dicek sebelum dan sesudah; bila beda otomatis dibatalkan. Lanjut?',ui.ButtonSet.YES_NO)!==ui.Button.YES)return;
  var r=bpKunci_(bpUpgrade_);
  ui.alert(r.sudah?'Selesai (sheet diperiksa ulang).':'Upgrade berhasil: '+r.impor+' upah disalin. Total upah individual sama persis ('+r.total+'). Isi/ubah upah di sheet BOBOT PEK; material di sheet MATERIAL dan AHS.');
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
  if(ui.alert('ROLLBACK','Kembalikan rumus upah NILAI PROGRES ke kondisi sebelum upgrade (dari sheet '+BP.AR+'). Sheet BOBOT PEK, AHS dan MATERIAL tidak dihapus. Lanjut?',ui.ButtonSet.YES_NO)!==ui.Button.YES)return;
  bpKunci_(bpRollback_);
  ui.alert('Rollback selesai.');
}
function tarifBersih(){
  var ui=SpreadsheetApp.getUi();
  if(ui.alert('Bersihkan sisa versi lama','HAPUS PERMANEN: sheet TARIF_PEKERJAAN, BACKUP_TARIF_*, ARSIP_UPGRADE_V3, kolom ID di RAB UPDATE (Q), judul lama di NILAI PROGRES (AW:AX), dan isi tabel TARIF lama di REKAP. Setelah ini ROLLBACK tidak bisa lagi. Lakukan hanya bila BOBOT PEK sudah dipakai dan hasilnya benar. Lanjut?',ui.ButtonSet.YES_NO)!==ui.Button.YES)return;
  var o=bpKunci_(bpBersih_);
  ui.alert(o.length?'Dibersihkan:\n- '+o.join('\n- '):'Tidak ada sisa lama yang perlu dibersihkan.');
}
function hargaAtur(){
  var ui=SpreadsheetApp.getUi(),P=PropertiesService.getScriptProperties();
  var a=ui.prompt('Atur koneksi harga online','Pilih mesin pencari:\n1 = Claude (pencarian web + ekstraksi harga, butuh API key Anthropic)\n2 = Google Programmable Search (butuh API key + ID mesin/cx)\nKetik 1 atau 2:',ui.ButtonSet.OK_CANCEL);
  if(a.getSelectedButton()!==ui.Button.OK)return;
  var p=String(a.getResponseText()).trim();
  if(p!=='1'&&p!=='2'){ui.alert('Pilihan tidak dikenal. Tidak ada yang diubah.');return;}
  var k=ui.prompt('API key','Tempel API key (disimpan di Script Properties, bukan di sheet):',ui.ButtonSet.OK_CANCEL);
  if(k.getSelectedButton()!==ui.Button.OK||!String(k.getResponseText()).trim()){ui.alert('Dibatalkan. Tidak ada yang diubah.');return;}
  P.setProperty('HARGA_PROVIDER',p==='1'?'CLAUDE':'GOOGLE');
  P.setProperty('HARGA_API_KEY',String(k.getResponseText()).trim());
  if(p==='2'){
    var c=ui.prompt('ID mesin pencari (cx)','Tempel cx Programmable Search Engine:',ui.ButtonSet.OK_CANCEL);
    if(c.getSelectedButton()!==ui.Button.OK||!String(c.getResponseText()).trim()){ui.alert('cx kosong. Koneksi Google belum lengkap.');return;}
    P.setProperty('HARGA_CSE_CX',String(c.getResponseText()).trim());
  }
  ui.alert('Koneksi disimpan ('+(p==='1'?'Claude':'Google')+'). Isi YA di kolom "CARI ONLINE?" sheet MATERIAL, lalu menu Material > Sync harga online.');
}
function hargaSync(){
  var ui=SpreadsheetApp.getUi();
  if(ui.alert('Sync harga online','Mencari harga online untuk material bertanda YA (maks 10 per proses, yang paling lama belum dicek lebih dulu). Hasil masuk ke kolom USULAN; harga master TIDAK berubah sampai Anda memilih "Terapkan usulan harga". Lanjut?',ui.ButtonSet.YES_NO)!==ui.Button.YES)return;
  var r=bpKunci_(function(){return hargaSync_(10);});
  ui.alert('Selesai: '+r.dicek+' dicek, '+r.ketemu+' harga ditemukan, '+r.gagal+' gagal'+(r.sisa>0?', '+r.sisa+' belum diproses (jalankan lagi).':'.')+(r.galat?'\nBerhenti: '+r.galat:'')+'\nPeriksa kolom SUMBER dan SATUAN sebelum menerapkan.');
}
function hargaTerapkan(){
  var ui=SpreadsheetApp.getUi(),L=hargaDaftarUsulan_();
  if(!L.length){ui.alert('Tidak ada usulan harga yang berbeda dari harga master.');return;}
  var t=L.slice(0,12).map(function(x){return x.nama+' ('+x.sat+'): '+Math.round(x.lama).toLocaleString('id-ID')+' -> '+Math.round(x.baru).toLocaleString('id-ID')+(x.lama>0?' ('+(x.baru/x.lama>=1?'+':'')+((x.baru/x.lama-1)*100).toFixed(1)+'%)':'');}).join('\n');
  if(ui.alert('Terapkan usulan harga',L.length+' harga akan diganti:\n'+t+(L.length>12?'\n...':'')+'\n\nHapus dulu sel USULAN yang tidak ingin dipakai. Lanjut?',ui.ButtonSet.YES_NO)!==ui.Button.YES)return;
  var n=bpKunci_(hargaTerapkan_);ui.alert(n+' harga diterapkan. AHS dan RAP ikut berubah otomatis.');
}

/* ====================== FUNGSI LAMA YANG DIGANTI ====================== */
/* GANTI onEdit lama dengan ini. */
function onEdit(e){
  try{
    if(!e||!e.range)return;
    var sh=e.range.getSheet(),nm=sh.getName(),r1=e.range.getRow(),r2=e.range.getLastRow(),c1=e.range.getColumn(),c2=e.range.getLastColumn();
    if(nm===BP.SH){bpStempel_(sh,r1,r2,c1,c2);return;}
    if(nm===MT.SH){matStempel_(sh,r1,r2,c1,c2);return;}
    if(nm!=='REKAP PELAKSANA (Rp)'||c1>8||c2<8)return;
    if(bpAktif_())bpRekapEdit_(sh,r1,r2,c1,c2);
    else SpreadsheetApp.getActive().toast('Upah satuan kini diisi di sheet BOBOT PEK. Jalankan menu Tarif > Upgrade / perbarui dulu.','Upah',10);
  }catch(x){SpreadsheetApp.getActive().toast('Gagal: '+x.message,'Upah',10);}
}
/* GANTI onOpen lama dengan ini (menu Sinkron tetap; ditambah menu Tarif dan Material). */
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
    .addItem('Upgrade / perbarui BOBOT PEK + AHS + MATERIAL (aman diulang)','tarifUpgrade')
    .addSeparator()
    .addItem('Bersihkan sisa versi lama (hapus permanen)','tarifBersih')
    .addItem('ROLLBACK upgrade','tarifRollback')
    .addToUi();
  ui.createMenu('Material')
    .addItem('Sync harga online','hargaSync')
    .addItem('Terapkan usulan harga','hargaTerapkan')
    .addItem('Atur koneksi harga online','hargaAtur')
    .addToUi();
}
