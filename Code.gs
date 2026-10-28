/** Annie Lytics report filing. No AI, email writes, sharing changes, or web endpoint. */
const PARSER_VERSION='static-v4';
const CFG = Object.freeze({
  account:'annie.lytics2026@gmail.com', root:'1H7gpTsWOd2MJkPCo0_6-zk8my8fCD6GZ',
  review:'1CFqyFWnmtTt7Qxc5tSuPrltzP-BwpNZk', system:'1LGBzcZWeh750ArvyTMcvey6-rHFmyxX7',
  timezone:'America/New_York', hours:[4,12,19], maxMs:240000,
  projectUrl:'https://script.google.com/home/projects/1N2GSrAEUsJ4xBr9ixVr6Hf-fgciOO3RvGKnhYaiZsWHZVfPIw-Lp8ZYR/edit',
  parents:{'ClubOS':{'Herndon':'15qyxU5djrnHP5QhfQSJx3dssQ3Vp5HIg','Chantilly':'1nvNwKLmj1LkhbZifnh0UZTc5PdI3FEC_','Both Clubs':'1fsxGrRrRS_TYF4s0q6XS6m6IMCXPIDAD'},'ABC Financial':{'Herndon':'1Rt6H9UQoMfKNNCITvNqelRJg1Eq0izJv','Chantilly':'1bOVZl56MrTqkq_Id63KIX3yDZoaFgwgU','Both Clubs':'1r9CwraqyyKvppI6mXsUzUP564fNyo2Gc'}}
});
function setupControl() {
  const lock=LockService.getScriptLock(); lock.waitLock(10000);
  try { assertAccount_(); const ss=control_(); console.log('Control sheet: '+ss.getUrl()); }
  finally {lock.releaseLock();}
}
function syncNow() { runSync_(null); }
function scheduledSync() { runSync_(null); }
function previewTest() {
  const ids=JSON.parse(PropertiesService.getScriptProperties().getProperty('TEST_MESSAGE_IDS')||'[]');
  if(ids.length!==3) throw Error('Set exactly three approved test message IDs first.');
  runSync_(ids);
}
function enableSchedule() {
  const p=PropertiesService.getScriptProperties();
  if(p.getProperty('LIVE_TESTS_PASSED')!=='yes') throw Error('Live single-club, both-club, unknown and repeat tests must pass first.');
  assertAccount_();
  const lock=LockService.getScriptLock();lock.waitLock(10000);
  try {
    const old=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='scheduledSync');
    // Reconcile only this script's filing triggers; retries after partial creation are safe.
    old.forEach(t=>ScriptApp.deleteTrigger(t));
    CFG.hours.forEach(h=>ScriptApp.newTrigger('scheduledSync').timeBased().everyDays(1).atHour(h).nearMinute(0).inTimezone(CFG.timezone).create());
    console.log('Created 3 daily triggers around 04:00, 12:00 and 19:00 America/New_York.');
  } finally {lock.releaseLock();}
}
function assertAccount_() {
  validateFolder_(CFG.root);
  const p=api_('https://gmail.googleapis.com/gmail/v1/users/me/profile');
  if(p.emailAddress!==CFG.account) throw Error('Wrong Gmail account; stopped.');
  const f=api_('https://www.googleapis.com/drive/v3/files/'+CFG.root+'?fields=id,name,owners(emailAddress),permissions(type,role,emailAddress),trashed');
  if(f.trashed || !(f.owners||[]).some(o=>o.emailAddress===CFG.account)) throw Error('Reports root ownership mismatch.');
}
function control_() {
  const props=PropertiesService.getScriptProperties();let id=props.getProperty('CONTROL_ID');
  const ss=id?SpreadsheetApp.openById(id):SpreadsheetApp.create('Anytime Fitness Reports — Sync & Index');
  if(id){if(props.getProperty('CONTROL_READY')==='yes'){validateControl_(id);return ss;}validatePendingControl_(id);}
  // Store ID immediately so a later setup error resumes this same sheet.
  props.setProperty('CONTROL_ID',ss.getId());
  validateFolder_(CFG.root);
  DriveApp.getFileById(ss.getId()).moveTo(DriveApp.getFolderById(CFG.root));
  validateControl_(ss.getId());
  const home=ss.getSheets()[0];home.setName('Sync');
  home.getRange(1,1,8,2).setValues([
    ['Report filing','Setup / live verification pending'],
    ['Sync now',CFG.projectUrl],
    ['How to sync','Open script above, select syncNow, then Run.'],
    ['Schedule','04:00, 12:00, 19:00 America/New_York; approximate windows; not yet enabled'],
    ['Last completed run','Not run'],['Report folder','https://drive.google.com/drive/folders/'+CFG.root],
    ['Agent chat sync','Pending authenticated workspace integration'],
    ['Version policy','All versions retained. Conflicting content for one period requires review.']]);
  home.setColumnWidth(1,190);home.setColumnWidth(2,670);home.getDataRange().setWrap(true);
  const ledger=(ss.getSheetByName('Receipts')||ss.insertSheet('Receipts'));if(ledger.getLastRow()===0)ledger.appendRow(['Source key','State','Drive file ID','Record JSON']);ledger.setFrozenRows(1);
  const latest=(ss.getSheetByName('Latest by Period')||ss.insertSheet('Latest by Period'));if(latest.getLastRow()===0)latest.appendRow(['Provider','Club','Report type','Period start','Period end','Latest received','Files','Version conflict']);latest.setFrozenRows(1);
  const runs=(ss.getSheetByName('Runs')||ss.insertSheet('Runs'));if(runs.getLastRow()===0)runs.appendRow(['Started UTC','Finished UTC','Processed','State','Details']);runs.setFrozenRows(1);
  ss.setSpreadsheetTimeZone(CFG.timezone);props.setProperty('CONTROL_READY','yes');return ss;
}
function runSync_(testIds) {
  const lock=LockService.getScriptLock();if(!lock.tryLock(1000)) {console.log('Another sync is running.');return {state:'BUSY'};}
  const start=Date.now();let processed=0,ss,activeKey=null,recordOutcome_,stage='STARTUP_ACCOUNT';
  const metrics={checked:0,skipped:0,imported:0,reclassified:0,excluded:0,gmailDownloads:0,driveDownloads:0},folderCache={};
  const detail_=text=>text+'; run counters '+JSON.stringify(metrics);
  try {
    assertAccount_();stage='STARTUP_CONTROL';ss=control_();stage='PROCESSING';
    const props=PropertiesService.getScriptProperties();
    if(!testIds && props.getProperty('IMPORT_APPROVED')!=='yes') throw Error('Import not activated: finish live tests then set IMPORT_APPROVED=yes.');
    const sheet=ss.getSheetByName('Receipts'), rows=sheet.getDataRange().getValues(), index={};
    for(let i=1;i<rows.length;i++)if(rows[i][0]){const r=JSON.parse(rows[i][3]);validateStoredReceipt_(r);if(index[rows[i][0]]||rows[i][0]!==r.sourceKey||rows[i][1]!==r.state||rows[i][2]!==r.fileId)throw Error('Ledger mismatch');index[r.sourceKey]={row:i+1,record:r};}
    const audit=ss.getSheetByName('Outcomes')||ss.insertSheet('Outcomes');
    if(!audit.getLastRow())audit.appendRow(['Source key','Outcome','Reason']);
    const outcomeRows=audit.getDataRange().getValues(),outcomes={};
    outcomeRows.slice(1).forEach((r,i)=>{if(!/^[A-Za-z0-9_-]+(?::[0-9.]*)?$/.test(r[0])||outcomes[r[0]]||!['IMPORTED','EXCLUDED','FAILED','REVIEW_REQUIRED','RESOLVED'].includes(r[1]))throw Error('Invalid outcome ledger');outcomes[r[0]]={row:i+2,status:r[1],reason:r[2]};});
    const history=ss.getSheetByName('Outcome History')||ss.insertSheet('Outcome History');if(!history.getLastRow())history.appendRow(['UTC','Source key','Previous outcome','New outcome','Reason']);
    recordOutcome_=(key,status,reason)=>{
      const old=outcomes[key],values=[key,status,reason];
      if(old&&old.status===status&&old.reason===reason)return;
      if(!old||old.status!==status||old.reason!==reason)history.appendRow(literalGrid_([[new Date().toISOString(),key,old?old.status:'',status,reason]])[0]);
      if(old)audit.getRange(old.row,1,1,3).setValues([values]);else audit.appendRow(values);
      outcomes[key]={row:old?old.row:audit.getLastRow(),status,reason};SpreadsheetApp.flush();
    };
    let cursor=testIds?{ids:testIds,next:null,index:0,part:0}:JSON.parse(props.getProperty('SCAN_CURSOR')||'null'),done=false;
    function saveCursor_(){if(!testIds){props.setProperty('SCAN_CURSOR',JSON.stringify(cursor));props.deleteProperty('PAGE_TOKEN');}}
    function advancePart_(){cursor.part++;saveCursor_();}
    function resolveMessage_(){const id=cursor.ids[cursor.index];if(outcomes[id]&&outcomes[id].status==='FAILED')recordOutcome_(id,'RESOLVED','Message retry succeeded; attachment outcomes recorded separately');}
    if(cursor&&(!Array.isArray(cursor.ids)||cursor.ids.length>20||cursor.ids.some(id=>! /^[A-Za-z0-9_-]+$/.test(id))||!Number.isInteger(cursor.index)||cursor.index<0||cursor.index>cursor.ids.length||!Number.isInteger(cursor.part)||cursor.part<0||cursor.part>1000||(cursor.next!==null&&typeof cursor.next!=='string')))throw Error('Invalid scan cursor');
    do {
      if(!cursor){
       const page=props.getProperty('PAGE_TOKEN');
       const response=api_('https://gmail.googleapis.com/gmail/v1/users/me/messages?maxResults=20&q='+encodeURIComponent('has:attachment -in:trash -in:spam')+(page?'&pageToken='+encodeURIComponent(page):''));
       cursor={ids:(response.messages||[]).map(m=>m.id),next:response.nextPageToken||null,index:0,part:0};saveCursor_();
      }
      for(;cursor.index<cursor.ids.length;resolveMessage_(),cursor.index++,cursor.part=0,saveCursor_()) {
        const ref={id:cursor.ids[cursor.index]};activeKey=ref.id;
        if(Date.now()-start>CFG.maxMs) { rebuildLatest_(ss,index);logRun_(ss,start,processed,'PAUSED',detail_('Resume from saved message/attachment cursor'));return {state:'PAUSED'}; }
        const msg=api_('https://gmail.googleapis.com/gmail/v1/users/me/messages/'+ref.id+'?format=full');
        const headers={};(msg.payload.headers||[]).forEach(h=>headers[h.name.toLowerCase()]=h.value);
        const subject=headers.subject||'',sender=(headers.from||'').toLowerCase();
        const trusted=trustedMessage_(msg.payload.headers||[]);
        const parts=[];walkParts_(msg.payload,parts);
        if(parts.length>1000)throw Error('MIME part bound');
        while(cursor.part<parts.length) {
          const part=parts[cursor.part];activeKey=msg.id+':'+part.partId;
          if(Date.now()-start>CFG.maxMs||processed>=20){rebuildLatest_(ss,index);logRun_(ss,start,processed,'PAUSED',detail_('Work limit reached; attachment cursor saved'));return {state:'PAUSED'};}
          metrics.checked++;
          if(!trusted){metrics.excluded++;recordOutcome_(activeKey,'EXCLUDED','Sender authentication/provenance unavailable; original remains in Gmail');advancePart_();continue;}
          const name=part.filename||'';
          if(name.length>512||!/\.(xls|xlsx|csv|pdf)$/i.test(name)){metrics.excluded++;recordOutcome_(activeKey,'EXCLUDED','Unsupported filename or unrelated attachment format');advancePart_();continue;}
          const known=/^(?:Scheduled|DailyScheduled|CommissionReport|Full-Usage|daily_|af_attrition|active_mbrs|annual_fee|renewal_audit|skip_months)/i.test(name);
          if(!known && !(trusted && /report|billing|reconciliation/i.test(subject))){metrics.excluded++;recordOutcome_(activeKey,'EXCLUDED','Not identified as a report attachment');advancePart_();continue;}
          const sourceKey=msg.id+':'+part.partId;
          if(!part.body||!Number.isFinite(part.body.size)||part.body.size>5*1024*1024){metrics.excluded++;recordOutcome_(sourceKey,'EXCLUDED','Attachment missing size or exceeds 5 MiB; original remains in Gmail');advancePart_();continue;}
          if(Date.now()-start>CFG.maxMs||processed>=20){rebuildLatest_(ss,index);logRun_(ss,start,processed,'PAUSED',detail_('Work limit reached; attachment cursor saved'));return {state:'PAUSED'};}
          let entry=index[sourceKey],rec=entry?entry.record:null;
          const wasCommitted=!!rec&&rec.state==='COMMITTED';
          const wantsReclassification=wasCommitted&&props.getProperty('RECLASSIFY_APPROVED_VERSION')===PARSER_VERSION&&rec.status!=='READY'&&rec.attemptVersion!==PARSER_VERSION&&/\.(xlsx?|pdf)$/i.test(name);
          // Gmail message+MIME-part identities are immutable; never use an attachment fetch ID as identity.
          // The fresh Drive server checksum must match the previously byte-verified receipt.
          // Missing or differing checksums fall through to the original full-byte verification.
          let prefetchedFile=null;
          if(wasCommitted&&!wantsReclassification){
            validateReceipt_(rec,{sourceKey,messageId:msg.id,mimePartId:part.partId,sha256:rec.sha256,filename:name});
            prefetchedFile=getFileMaybe_(rec.fileId);
            if(prefetchedFile){
              validateOriginal_(prefetchedFile,rec,folderCache);
              if(prefetchedFile.parents[0]!==rec.destination)throw Error('Committed destination mismatch');
              if(/^[a-f0-9]{64}$/.test(prefetchedFile.sha256Checksum||'')&&prefetchedFile.sha256Checksum===rec.sha256&&String(part.body.size)===String(prefetchedFile.size)){
                recordOutcome_(sourceKey,rec.status==='READY'?'IMPORTED':'REVIEW_REQUIRED',rec.status==='READY'?'Filed original':'Original saved; classification requires review');
                metrics.skipped++;advancePart_();continue;
              }
            }
          }
          if(part.body.attachmentId)metrics.gmailDownloads++;
          const encoded=part.body.attachmentId?api_('https://gmail.googleapis.com/gmail/v1/users/me/messages/'+msg.id+'/attachments/'+encodeURIComponent(part.body.attachmentId)).data:part.body.data;
          if(!encoded)throw Error('Missing bytes for '+sourceKey);
          const bytes=Utilities.base64DecodeWebSafe(encoded);if(bytes.length>5*1024*1024){metrics.excluded++;recordOutcome_(sourceKey,'EXCLUDED','Decoded attachment exceeds 5 MiB; original remains in Gmail');advancePart_();continue;}const sha=sha_(bytes);
          const blob=Utilities.newBlob(bytes,'application/octet-stream',name);
          if(rec)validateReceipt_(rec,{sourceKey,messageId:msg.id,mimePartId:part.partId,attachmentId:part.body.attachmentId||null,sha256:sha,filename:name});
          if(!rec) {
            const fileId=api_('https://www.googleapis.com/drive/v3/files/generateIds?count=1&space=drive&type=files').ids[0];
            rec={sourceKey,account:CFG.account,messageId:msg.id,attachmentId:part.body.attachmentId||null,mimePartId:part.partId,
              filename:name,filenameSha256:sha_(Utilities.newBlob(name).getBytes()),sha256:sha,receivedAt:new Date(Number(msg.internalDate)).toISOString(),
              savedAt:null,reportDate:null,periodStart:null,periodEnd:null,provider:null,club:null,type:null,state:'PREPARED',fileId,
              status:'NEEDS_REVIEW',reason:'Parsing pending',destination:CFG.review};
            sheet.appendRow([sourceKey,'PREPARED',fileId,JSON.stringify(rec)]);SpreadsheetApp.flush();entry={row:sheet.getLastRow(),record:rec};index[sourceKey]=entry;
          }
          validateFolder_(CFG.review);
          let file=prefetchedFile||getFileMaybe_(rec.fileId);
          if(file)validateOriginal_(file,rec);
          if(!file) {
            multipart_({id:rec.fileId,name,parents:[CFG.review],appProperties:{reportSource:sourceKey,sha256:sha}},blob);
            file=api_('https://www.googleapis.com/drive/v3/files/'+rec.fileId+'?fields=id,parents,sha256Checksum,createdTime,owners(emailAddress),permissions(type,role,emailAddress),trashed,appProperties');
          }
          metrics.driveDownloads++;
          const saved=googleFetch_('https://www.googleapis.com/drive/v3/files/'+rec.fileId+'?alt=media',{headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()}}).getBlob().getBytes();
          if(sha_(saved)!==sha)throw Error('Saved attachment hash mismatch: '+sourceKey);
          rec.savedAt=file.createdTime||new Date().toISOString();
          if(rec.state==='COMMITTED'){
            validateFolder_(rec.destination);if(file.parents.length!==1||file.parents[0]!==rec.destination)throw Error('Committed destination mismatch');
            const reclassify=props.getProperty('RECLASSIFY_APPROVED_VERSION')===PARSER_VERSION&&rec.status!=='READY'&&rec.attemptVersion!==PARSER_VERSION&&/\.(xlsx?|pdf)$/i.test(name);
            if(!reclassify){metrics.skipped++;recordOutcome_(sourceKey,rec.status==='READY'?'IMPORTED':'REVIEW_REQUIRED',rec.status==='READY'?'Filed original':'Original saved; classification requires review');advancePart_();continue;}
            rec.previousClassification={status:rec.status,reason:rec.reason,destination:rec.destination};rec.state='PREPARED';writeReceipt_(sheet,entry,rec);
          }
          let evidence={status:'NEEDS_REVIEW',reason:'Unverified sender provenance; original only'};
          if(trusted) {
            evidence={status:'NEEDS_REVIEW',reason:'Static parser unavailable for this format; original only'};
            if(/\.xlsx$/i.test(name)) {
              try {evidence=parseAbcWorkbook(staticXlsx_(blob));}
              catch(e) {evidence={status:'NEEDS_REVIEW',reason:'Static workbook validation failed; original only'};}
            } else if(/\.pdf$/i.test(name)) {
              evidence=classifyKnownPdf_(sha);
            } else if(/\.xls$/i.test(name)) {
              try {evidence=parseClubOsStatic_(staticLegacyWorkbook_(blob));}
              catch(e) {evidence={status:'NEEDS_REVIEW',reason:'Static legacy workbook validation failed; original only'};}
            }
          }
          Object.assign(rec,evidence);rec.attemptVersion=PARSER_VERSION;
          const destination=['READY','ROUTED_DATE_REVIEW'].includes(rec.status)?reportFolder_(rec.provider,rec.club,rec.type):CFG.review;
          validateFolder_(destination);
          const current=api_('https://www.googleapis.com/drive/v3/files/'+rec.fileId+'?fields=id,parents,owners(emailAddress),permissions(type,role,emailAddress),trashed,appProperties');
          validateOriginal_(current,rec);
          if((current.parents||[]).join()!==destination) api_('https://www.googleapis.com/drive/v3/files/'+rec.fileId+'?addParents='+destination+'&removeParents='+encodeURIComponent((current.parents||[]).join(',')),{method:'patch',payload:'{}',contentType:'application/json'});
          const verified=api_('https://www.googleapis.com/drive/v3/files/'+rec.fileId+'?fields=id,parents');
          if((verified.parents||[]).indexOf(destination)<0)throw Error('Destination verification failed');
          rec.destination=destination;rec.state='COMMITTED';writeReceipt_(sheet,entry,rec);processed++;if(wasCommitted)metrics.reclassified++;else metrics.imported++;
          recordOutcome_(sourceKey,rec.status==='READY'?'IMPORTED':'REVIEW_REQUIRED',rec.status==='READY'?'Filed original':'Original saved; classification requires review');advancePart_();
        }
      }
      done=!cursor.next;
      if(!testIds){
       // Persist the exhausted page until the next page is durably loaded. A retry
       // skips its finished attachments and repeats only the read-only page fetch.
       if(done){props.deleteProperty('SCAN_CURSOR');props.deleteProperty('PAGE_TOKEN');}
       else{
        const response=api_('https://gmail.googleapis.com/gmail/v1/users/me/messages?maxResults=20&q='+encodeURIComponent('has:attachment -in:trash -in:spam')+'&pageToken='+encodeURIComponent(cursor.next));
        cursor={ids:(response.messages||[]).map(m=>m.id),next:response.nextPageToken||null,index:0,part:0};saveCursor_();
       }
      }
    } while(!done&&!testIds&&Date.now()-start<CFG.maxMs);
    const totals={IMPORTED:0,REVIEW_REQUIRED:0,EXCLUDED:0,FAILED:0,RESOLVED:0};Object.values(outcomes).forEach(o=>totals[o.status]++);
    rebuildLatest_(ss,index);logRun_(ss,start,processed,done?'COMPLETE':'PAUSED',detail_((testIds?'Approved test scope; ':'Mailbox scan; ')+'cumulative source outcomes '+JSON.stringify(totals)));
    ss.getSheetByName('Sync').getRange('B5').setValue(new Date().toISOString()+'; '+(done?'complete':'paused'));
    return {state:done?'COMPLETE':'PAUSED'};
  } catch(e) {const failure=stage+':'+safeFailureCode_(e);console.log('SYNC_FAILURE '+failure);if(recordOutcome_&&activeKey){try{recordOutcome_(activeKey,'FAILED','Run stopped at this source; retry required');}catch(ignored){}}if(ss)logRun_(ss,start,processed,'ERROR',detail_('Run failed: '+failure));throw Error('Report filing stopped ['+failure+']; no source files deleted');}
  finally {lock.releaseLock();}
}
function safeFailureCode_(e){
 const message=String(e&&e.message||'');
 const known={'Unexpected report access':'ACCESS_POLICY_MISMATCH','Cannot verify permissions':'ACCESS_UNVERIFIABLE','Wrong Gmail account; stopped.':'ACCOUNT_MISMATCH','Reports root ownership mismatch.':'ROOT_OWNERSHIP_MISMATCH','Invalid folder ownership':'FOLDER_OWNERSHIP_MISMATCH','Invalid control sheet':'CONTROL_METADATA_MISMATCH','Folder outside report root':'FOLDER_OUTSIDE_ROOT'};
 if(known[message])return known[message];
 const http=/^(?:Google API HTTP |Drive lookup failed HTTP )(400|401|403|404|409|429|500|502|503|504)\b/.exec(message);
 return http?'GOOGLE_HTTP_'+http[1]:'UNCLASSIFIED_ERROR';
}
function reportFolder_(provider,club,type) {
  const parent=CFG.parents[provider]&&CFG.parents[provider][club];if(!parent)throw Error('Unconfigured route');
  validateFolder_(parent);
  const folder=DriveApp.getFolderById(parent),matches=folder.getFoldersByName(type);
  if(matches.hasNext()){const f=matches.next();if(matches.hasNext())throw Error('Ambiguous destination folders');validateFolder_(f.getId());return f.getId();}
  return folder.createFolder(type).getId();
}
function rebuildLatest_(ss,index) {
  const groups={};Object.values(index).forEach(e=>{const r=e.record;if(r.state!=='COMMITTED'||r.status!=='READY')return;const k=JSON.stringify([r.provider,r.club,r.type,r.periodStart,r.periodEnd]);(groups[k]||(groups[k]=[])).push(r);});
  const rows=Object.keys(groups).sort().map(k=>{const rs=groups[k],hashes=new Set(rs.map(r=>r.sha256));return JSON.parse(k).concat([rs.map(r=>r.receivedAt).sort().pop(),rs.map(r=>'https://drive.google.com/file/d/'+r.fileId+'/view').join('\n'),hashes.size>1?'YES — all versions preserved':'No']);});
  const sh=ss.getSheetByName('Latest by Period');if(sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,8).clearContent();if(rows.length)sh.getRange(2,1,rows.length,8).setValues(literalGrid_(rows));

  const reviewRows=Object.values(index).map(e=>e.record).filter(r=>r.state==='COMMITTED'&&r.status!=='READY').map(r=>[r.sourceKey,r.filename,r.receivedAt,r.reason,'https://drive.google.com/file/d/'+r.fileId+'/view',r.sha256]);
  const reviewSheet=ss.getSheetByName('Needs Review')||ss.insertSheet('Needs Review');
  if(!reviewSheet.getLastRow())reviewSheet.appendRow(['Source key','Original filename','Received UTC','Review reason','Original file link','SHA256']);
  reviewSheet.setFrozenRows(1);
  if(reviewSheet.getLastRow()>1)reviewSheet.getRange(2,1,reviewSheet.getLastRow()-1,6).clearContent();
  if(reviewRows.length)reviewSheet.getRange(2,1,reviewRows.length,6).setValues(literalGrid_(reviewRows));
}
function writeReceipt_(sheet,entry,r){sheet.getRange(entry.row,1,1,4).setValues([[r.sourceKey,r.state,r.fileId,JSON.stringify(r)]]);SpreadsheetApp.flush();entry.record=r;}
function walkParts_(p,out){if(p.filename)out.push(p);(p.parts||[]).forEach(x=>walkParts_(x,out));}
function validDate_(s){return /^\d{4}-\d{2}-\d{2}$/.test(s)&&!isNaN(Date.parse(s))&&new Date(s+'T00:00:00Z').toISOString().slice(0,10)===s;}
function sha_(bytes){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,bytes).map(x=>('0'+(x&255).toString(16)).slice(-2)).join('');}
function logRun_(ss,start,count,state,details){ss.getSheetByName('Runs').appendRow(literalGrid_([[new Date(start).toISOString(),new Date().toISOString(),count,state,details]])[0]);console.log(state+': '+count+' attachments; '+details);}
function getFileMaybe_(id){const r=googleFetch_('https://www.googleapis.com/drive/v3/files/'+id+'?fields=id,parents,size,sha256Checksum,createdTime,owners(emailAddress),permissions(type,role,emailAddress),trashed,appProperties',{headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},muteHttpExceptions:true});if(r.getResponseCode()===404)return null;if(r.getResponseCode()!==200)throw Error('Drive lookup failed HTTP '+r.getResponseCode());return JSON.parse(r.getContentText());}
function api_(url,options){options=Object.assign({},options||{});options.headers={Authorization:'Bearer '+ScriptApp.getOAuthToken()};options.muteHttpExceptions=true;const r=googleFetch_(url,options),code=r.getResponseCode();if(code<200||code>=300)throw Error('Google API HTTP '+code+' at '+url.split('?')[0]+'; check API enablement/permissions/quota');const t=r.getContentText();return t?JSON.parse(t):{};}
function multipart_(metadata,blob){if(blob.getContentType()!=='application/octet-stream'||metadata.mimeType)throw Error('Native conversion prohibited');if(!metadata.parents||metadata.parents.length!==1||metadata.parents[0]!==CFG.review||Object.keys(metadata).some(k=>['id','name','parents','appProperties'].indexOf(k)<0))throw Error('Upload metadata outside report scope');validateFolder_(CFG.review);const boundary='annie_'+Utilities.getUuid();const a=Utilities.newBlob('--'+boundary+'\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n'+JSON.stringify(metadata)+'\r\n--'+boundary+'\r\nContent-Type: '+blob.getContentType()+'\r\n\r\n').getBytes();const b=Utilities.newBlob('\r\n--'+boundary+'--').getBytes();return api_('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,createdTime',{method:'post',contentType:'multipart/related; boundary='+boundary,payload:a.concat(blob.getBytes(),b)});}

var ABC_TYPES = {'Daily MTD Billing Statement':'MTD Billing','Daily Delinquency':'Delinquency','Daily New Accounts':'New Accounts','Monthly Billing Statement':'Monthly Billing','Annual Fee Audit':'Annual Fee Audit','Renewal Audit Details':'Renewal Audit','Account Reconciliation':'Account Reconciliation','Daily Account Adjustments':'Account Adjustments','End of Term Skip':'End of Term Skip','Daily Production Report':'Production','Active Member Detail':'Active Members','AF Attrition Summary':'Attrition'};
function abcNorm(v){return String(v==null?'':v).replace(/\s+/g,' ').trim();}
function abcReview(reason){return {status:'NEEDS_REVIEW',reason:reason};}
function abcDate(s){var m=/^(January|February|March|April|May|June|July|August|September|October|November|December) (\d{1,2}), (\d{4})$/.exec(s);if(!m)return null;var n=['January','February','March','April','May','June','July','August','September','October','November','December'].indexOf(m[1])+1;var iso=m[3]+'-'+('0'+n).slice(-2)+'-'+('0'+m[2]).slice(-2),d=new Date(iso+'T00:00:00Z');return !isNaN(d)&&d.toISOString().slice(0,10)===iso?iso:null;}
function abcPeriod(text,title){var date='(?:January|February|March|April|May|June|July|August|September|October|November|December) \\d{1,2}, \\d{4}';var re=new RegExp(title+'\\s*(?:As of )?('+date+')(?: to ('+date+'))?','g'),m,seen={};while((m=re.exec(text))){var a=abcDate(m[1]),b=abcDate(m[2]||m[1]);if(!a||!b||a>b)return null;seen[a+'|'+b]=true;}if(title==='Annual Fee Audit'){re=/Annual Fee Audit\s*([A-Z][a-z]+) (\d{4})\b/g;while((m=re.exec(text))){var first=abcDate(m[1]+' 1, '+m[2]);if(!first)return null;var d=new Date(first+'T00:00:00Z');d.setUTCMonth(d.getUTCMonth()+1);d.setUTCDate(0);seen[first+'|'+d.toISOString().slice(0,10)]=true;}}var keys=Object.keys(seen);return keys.length===1?keys[0].split('|'):null;}
function abcPair(id,name,clubs){var map={'5188':'Chantilly','05188':'Chantilly','50213':'Herndon'},club=map[abcNorm(id)];if(!club||abcNorm(name)!=='ANYTIME FITNESS '+club.toUpperCase())return false;clubs[club]=true;return true;}
function abcResult(title,period,clubs){var names=Object.keys(clubs);if(!period||!names.length)return abcReview('Missing or conflicting ABC reporting period / club evidence');return {status:'READY',reason:'',provider:'ABC Financial',club:names.length===2?'Both Clubs':names[0],type:ABC_TYPES[title],periodStart:period[0],periodEnd:period[1],reportDate:period[0]===period[1]?period[0]:null,parserVersion:'abc-content-v1'};}
function parseAbcWorkbook(sheets){if(sheets.length!==1)return abcReview('Unverified ABC workbook sheet structure');var rows=sheets[0].rows.map(function(r){return r.map(abcNorm);}),title=rows[1]&&rows[1][0],clubs={};if(['Daily MTD Billing Statement','Daily Delinquency','Daily New Accounts'].indexOf(title)<0)return abcReview('Unverified ABC workbook title');var period=abcPeriod(title+' '+(rows[2]||[]).join(' '),title);if(title==='Daily MTD Billing Statement'){if(!/^Includes clubs from ANYTIME$/.test((rows[0]||[])[0]))return abcReview('Unverified billing header');var ids=(rows[5]||[]).filter(function(v){return !!v;});if(!ids.length)return abcReview('No billing club columns');for(var i=0;i<ids.length;i++){if(ids[i]==='5188')clubs.Chantilly=true;else if(ids[i]==='50213')clubs.Herndon=true;else return abcReview('Unknown billing club column');}}else{if(!rows[4]||rows[4][0]!=='Club Name'||rows[4][1]!=='Club Number')return abcReview('Unverified ABC club columns');for(var j=5;j<rows.length;j++){var name=rows[j][0],id=rows[j][1];if(/^ANYTIME FITNESS\b/.test(name)||/^\d+$/.test(id)){if(!abcPair(id,name,clubs))return abcReview('Unknown or conflicting ABC club name and number');}else if(name&&!/^(?:Total|Grand Total|Page|Report|Membership Types Selected:All Membership Types$|Run in Schedule by [A-Za-z0-9_]+ on [A-Za-z]+ \d{1,2}, \d{4} at \d{2}:\d{2}:\d{2} [AP]M Central Time$)/i.test(name))return abcReview('Unrecognized ABC club row');}}return abcResult(title,period,clubs);}
function parseAbcPdfText(raw){var text=abcNorm(raw),titles=Object.keys(ABC_TYPES).filter(function(t){return text.indexOf(t)>=0;});if(titles.length!==1)return abcReview('Missing or multiple ABC report titles');var title=titles[0],period=abcPeriod(text,title),clubs={},m;var before=/\b(\d{4,6})\s*-\s*(ANYTIME FITNESS [A-Z]+)\b/g;while((m=before.exec(text)))if(!abcPair(m[1],m[2],clubs))return abcReview('Unknown or conflicting ABC club header');var after=/(ANYTIME FITNESS [A-Z]+)\s+(\d{4,6})\b/g;while((m=after.exec(text)))if(!abcPair(m[2],m[1],clubs))return abcReview('Unknown or conflicting ABC club row');return abcResult(title,period,clubs);}

function googleFetch_(url,options){
 options=Object.assign({},options||{});const method=(options.method||'get').toLowerCase();
 const gmail=/^https:\/\/gmail\.googleapis\.com\/gmail\/v1\/users\/me\/(?:profile|messages(?:\/[A-Za-z0-9_-]+(?:\/attachments\/[A-Za-z0-9_%=-]+)?)?)(?:\?[^#]*)?$/.test(url);
 const drive=/^https:\/\/www\.googleapis\.com\/drive\/v3\/files\/[A-Za-z0-9_-]+(?:\/export)?(?:\?[^#]*)?$/.test(url);
 const upload=url==='https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,createdTime';
 const move=/^https:\/\/www\.googleapis\.com\/drive\/v3\/files\/[A-Za-z0-9_-]+\?addParents=[A-Za-z0-9_-]+&removeParents=[A-Za-z0-9_%,-]*$/.test(url);
 if(!((method==='get'&&(gmail||drive))||(method==='post'&&upload)||(method==='patch'&&move&&options.payload==='{}')))throw Error('Blocked unapproved API endpoint or operation');
 options.headers={Authorization:'Bearer '+ScriptApp.getOAuthToken()};options.muteHttpExceptions=true;options.followRedirects=false;
 return UrlFetchApp.fetch(url,options);
}

// Fail closed before interpreting untrusted content. Only Gmail's leading
// Authentication-Results entry for mx.google.com is accepted; later copies ignored.
function trustedMessage_(headers){
 const from=headers.filter(h=>h.name.toLowerCase()==='from');
 if(from.length!==1)return false;
 const value=from[0].value.trim().toLowerCase(),m=/^(?:[^<>\r\n]*<)?([a-z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-z0-9.-]+)>?$/.exec(value);
 if(!m||['annie.lytics@anytimefitness.com','reports@club-os.com','reportcaster@abcfitness.com'].indexOf(m[1])<0)return false;
 const auth=headers.find(h=>h.name.toLowerCase()==='authentication-results');
 if(!auth||!/^mx\.google\.com\s*;/i.test(auth.value))return false;
 const domain=m[1].split('@')[1],safe=domain.replace(/\./g,'\\.');
 return new RegExp('(?:^|;)\\s*dmarc=pass\\b[^;]*\\bheader\\.from='+safe+'(?:[;\\s]|$)','i').test(auth.value) &&
   new RegExp('(?:^|;)\\s*dkim=pass\\b[^;]*\\bheader\\.(?:i=@|d=)'+safe+'(?:[;\\s]|$)','i').test(auth.value);
}
function validateReceipt_(r,source){
 for(const k of ['sourceKey','messageId','mimePartId','sha256','filename'])if(r[k]!==source[k])throw Error('Receipt source mismatch');
 if(!['NEEDS_REVIEW','READY','ROUTED_DATE_REVIEW'].includes(r.status)||r.filename.length>512||!/^\d{4}-\d{2}-\d{2}T/.test(r.receivedAt))throw Error('Invalid receipt schema');
 if(r.account!==CFG.account||!/^[A-Za-z0-9_-]{10,}$/.test(r.fileId)||!['PREPARED','COMMITTED'].includes(r.state))throw Error('Invalid receipt');
 if(r.filenameSha256!==sha_(Utilities.newBlob(source.filename).getBytes()))throw Error('Receipt name mismatch');
}
function validateFolder_(id,cache){
 let current=id;const seen=new Set(),checked=[];
 // Only read-only committed-file skips use this short, per-run ancestry cache.
 // File ownership, permissions, provenance, parent and server checksum are always fetched afresh.
 const finish=stamp=>{if(cache){const at=Math.min(Date.now(),stamp||Date.now());checked.forEach(k=>cache[k]=at);}return;};
 for(let depth=0;depth<10;depth++){
  if(!/^[A-Za-z0-9_-]+$/.test(current)||seen.has(current))throw Error('Invalid folder ancestry');seen.add(current);
  if(cache&&cache[current]&&Date.now()-cache[current]<30000)return finish(cache[current]);
  checked.push(current);
  const f=api_('https://www.googleapis.com/drive/v3/files/'+current+'?fields=id,mimeType,parents,owners(emailAddress),permissions(type,role,emailAddress),trashed');
  validateSharing_(f);
  if(f.trashed||f.mimeType!=='application/vnd.google-apps.folder'||!(f.owners||[]).some(o=>o.emailAddress===CFG.account))throw Error('Invalid folder ownership');
  if(current===CFG.root)return finish();
  if(!f.parents||f.parents.length!==1)throw Error('Folder outside report root');current=f.parents[0];
 }
 throw Error('Folder ancestry too deep');
}
function validateOriginal_(f,r,cache){
 validateSharing_(f);
 if(f.id!==r.fileId||f.trashed||!(f.owners||[]).some(o=>o.emailAddress===CFG.account)||!f.appProperties||f.appProperties.reportSource!==r.sourceKey||f.appProperties.sha256!==r.sha256)throw Error('Original provenance mismatch');
 if(!f.parents||f.parents.length!==1)throw Error('Original ancestry mismatch');validateFolder_(f.parents[0],cache);
}
function validateControl_(id){
 const f=api_('https://www.googleapis.com/drive/v3/files/'+id+'?fields=id,mimeType,parents,owners(emailAddress),permissions(type,role,emailAddress),trashed');
 validateSharing_(f);
 if(f.trashed||f.mimeType!=='application/vnd.google-apps.spreadsheet'||!(f.owners||[]).some(o=>o.emailAddress===CFG.account)||!f.parents||f.parents.length!==1||f.parents[0]!==CFG.root)throw Error('Invalid control sheet');
 validateFolder_(CFG.root);
}
function zipBound_(bytes){
 if(bytes.length>5*1024*1024)throw Error('Archive too large');
 const u=i=>bytes[i]&255,le=(i,n)=>{let v=0;for(let j=0;j<n;j++)v+=u(i+j)*Math.pow(256,j);return v;};
 let e=-1;for(let i=bytes.length-22;i>=Math.max(0,bytes.length-65557);i--)if(le(i,4)===0x06054b50){e=i;break;}
 if(e<0||le(e+4,2)!==0||le(e+6,2)!==0)throw Error('Unsupported archive');
 const count=le(e+10,2),size=le(e+12,4),offset=le(e+16,4);
 if(count>500||count===65535||offset+size>e)throw Error('Archive bounds');
 let pos=offset,total=0;
 for(let k=0;k<count;k++){
  if(pos+46>e||le(pos,4)!==0x02014b50||le(pos+8,2)&1)throw Error('Invalid archive entry');
  total+=le(pos+24,4);if(total>20*1024*1024)throw Error('Expanded archive too large');
  pos+=46+le(pos+28,2)+le(pos+30,2)+le(pos+32,2);
 }
 if(pos!==offset+size)throw Error('Invalid central directory');
}
function staticXlsx_(blob){
 zipBound_(blob.getBytes());const entries=unzipBounded_(blob),files={};
 entries.forEach(b=>{const n=b.getName();if(!/^[A-Za-z0-9_./\[\]-]+$/.test(n)||n.includes('..')||files[n])throw Error('Unsafe archive path');files[n]=b.getDataAsString();});
 const ns=XmlService.getNamespace('http://schemas.openxmlformats.org/spreadsheetml/2006/main');
 function xml(name){const t=files[name];if(!t||/<!DOCTYPE|<!ENTITY/i.test(t))throw Error('Unsupported XML');return XmlService.parse(t).getRootElement();}
 for(const n of Object.keys(files)){
  if(/externalLinks|vbaProject|embeddings|connections/i.test(n))throw Error('Active content blocked');
  if(/\.rels$/.test(n)) {const rr=xml(n);rr.getChildren('Relationship',rr.getNamespace()).forEach(r=>{const mode=r.getAttribute('TargetMode'),type=r.getAttribute('Type');if(mode&&mode.getValue()==='External'&&(!type||type.getValue()!=='http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'))throw Error('External data relationship blocked');});} // Hyperlinks are inert metadata: never followed or written to index.

  if(/^xl\/worksheets\/.*\.xml$/.test(n)&&/<(?:\w+:)?f(?:\s|>)/.test(files[n]))throw Error('Formula workbook blocked');
 }
 const strings=files['xl/sharedStrings.xml']?xml('xl/sharedStrings.xml').getChildren('si',ns).map(e=>{
  const t=e.getChild('t',ns);return t?t.getText():e.getChildren('r',ns).map(r=>r.getChildText('t',ns)||'').join('');
 }):[];
 const rel=xml('xl/_rels/workbook.xml.rels'),rn=rel.getNamespace(),targets={};
 rel.getChildren('Relationship',rn).forEach(r=>{targets[r.getAttribute('Id').getValue()]=r.getAttribute('Target').getValue();});
 const rid=XmlService.getNamespace('http://schemas.openxmlformats.org/officeDocument/2006/relationships');
 let cellCount=0;
 return xml('xl/workbook.xml').getChild('sheets',ns).getChildren('sheet',ns).map(s=>{
  const target=targets[s.getAttribute('id',rid).getValue()];if(!target||target.includes('..'))throw Error('Unsupported worksheet target');
  const path=target.startsWith('/xl/')?target.slice(1):'xl/'+target;
  if(!/^xl\/worksheets\/[A-Za-z0-9_-]+\.xml$/.test(path))throw Error('Unsupported worksheet path');
  const rows=[];xml(path).getChild('sheetData',ns).getChildren('row',ns).forEach(row=>{
   const ri=Number(row.getAttribute('r').getValue())-1;if(ri<0||ri>50000)throw Error('Row bounds');const values=[];
   row.getChildren('c',ns).forEach(c=>{
    if(++cellCount>200000||c.getChild('f',ns))throw Error('Cell/formula bounds');
    const ref=c.getAttribute('r').getValue(),match=/^([A-Z]{1,3})\d+$/.exec(ref);if(!match)throw Error('Cell reference');
    let col=0;for(const ch of match[1])col=col*26+ch.charCodeAt(0)-64;if(col>200)throw Error('Column bounds');
    const type=c.getAttribute('t'),v=c.getChildText('v',ns)||'';let text=v;
    if(type&&type.getValue()==='s'){if(!/^\d+$/.test(v)||strings[Number(v)]===undefined)throw Error('Shared string index');text=strings[Number(v)];}
    else if(type&&type.getValue()==='inlineStr'){const is=c.getChild('is',ns);text=is?(is.getChildText('t',ns)||is.getChildren('r',ns).map(r=>r.getChildText('t',ns)||'').join('')):'';}
    values[col-1]=text;
   });rows[ri]=Array.from({length:values.length},(_,i)=>values[i]||'');
  });return {name:s.getAttribute('name').getValue(),rows:Array.from({length:rows.length},(_,i)=>rows[i]||[])};
 });
}

function literalGrid_(rows){return rows.map(row=>row.map(v=>typeof v==='string'&&/^[=+@\-\t\r]/.test(v)?"\'"+v:v));}

function validateSharing_(f){
 if(!Array.isArray(f.permissions)||!f.permissions.length)throw Error('Cannot verify permissions');
 for(const p of f.permissions)if(!(p.type==='user'&&((p.role==='owner'&&p.emailAddress===CFG.account)||(p.role==='reader'&&['brett@breakaway.us','trent@breakaway.us'].includes(p.emailAddress))||(p.role==='writer'&&p.emailAddress==='brett.mclaren@gmail.com'))))throw Error('Unexpected report access');
}

function validateStoredReceipt_(r){
 if(!r||r.account!==CFG.account||typeof r.messageId!=='string'||!/^[A-Za-z0-9_-]+$/.test(r.messageId)||typeof r.mimePartId!=='string'||!/^[0-9.]*$/.test(r.mimePartId)||r.sourceKey!==r.messageId+':'+r.mimePartId||typeof r.filename!=='string'||r.filename.length>512||!/^[a-f0-9]{64}$/.test(r.sha256)||!/^[a-f0-9]{64}$/.test(r.filenameSha256)||!/^[A-Za-z0-9_-]{10,}$/.test(r.fileId)||!['PREPARED','COMMITTED'].includes(r.state)||!['NEEDS_REVIEW','READY','ROUTED_DATE_REVIEW'].includes(r.status))throw Error('Invalid stored receipt');
 if(['READY','ROUTED_DATE_REVIEW'].includes(r.status)&&(!CFG.parents[r.provider]||!CFG.parents[r.provider][r.club]||typeof r.type!=='string'||!(r.provider==='ABC Financial'?Object.values(ABC_TYPES):Object.values(CLUBOS_TYPES)).includes(r.type)))throw Error('Invalid stored report classification');
 if(r.status==='READY'&&(!validDate_(r.periodStart)||!validDate_(r.periodEnd)||r.periodStart>r.periodEnd))throw Error('Invalid stored report period');
 if(r.status==='ROUTED_DATE_REVIEW'&&(r.periodStart!==null||r.periodEnd!==null||r.reportDate!==null||typeof r.reason!=='string'||!r.reason))throw Error('Invalid unresolved report period');
}

// Bounded DEFLATE decoder: output size is checked before every emitted byte.
function inflateBounded_(input,limit){
 let pos=0,bits=0,buf=0;const out=[];
 function read(n){while(bits<n){if(pos>=input.length)throw Error('Truncated deflate');buf|=(input[pos++]&255)<<bits;bits+=8;}const v=buf&((1<<n)-1);buf>>>=n;bits-=n;return v;}
 function tree(lengths){const count=Array(16).fill(0),next=Array(16).fill(0),table={};for(const n of lengths){if(n<0||n>15)throw Error('Huffman bound');if(n)count[n]++;}let code=0;for(let n=1;n<=15;n++){code=(code+count[n-1])<<1;next[n]=code;if(code+count[n]>(1<<n))throw Error('Oversubscribed Huffman');}lengths.forEach((n,s)=>{if(n){let c=next[n]++,rev=0;for(let j=0;j<n;j++){rev=(rev<<1)|(c&1);c>>>=1;}table[n+':'+rev]=s;}});return table;}
 function sym(t){let c=0;for(let n=1;n<=15;n++){c|=read(1)<<(n-1);const v=t[n+':'+c];if(v!==undefined)return v;}throw Error('Invalid Huffman symbol');}
 function put(v){if(out.length>=limit)throw Error('Inflate output bound');out.push(v);}
 const lb=[3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,35,43,51,59,67,83,99,115,131,163,195,227,258],le=[0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0],db=[1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,257,385,513,769,1025,1537,2049,3073,4097,6145,8193,12289,16385,24577],de=[0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,12,13,13];
 let final=0,blocks=0;
 do{
  if(++blocks>10000)throw Error('Deflate block bound');final=read(1);const kind=read(2);let lit,dist;
  if(kind===0){bits=0;buf=0;const len=read(16),inv=read(16);if((len^inv)!==65535)throw Error('Stored block length');for(let i=0;i<len;i++)put(read(8));continue;}
  if(kind===1){lit=tree(Array.from({length:288},(_,i)=>i<144?8:i<256?9:i<280?7:8));dist=tree(Array(32).fill(5));}
  else if(kind===2){const nl=read(5)+257,nd=read(5)+1,nc=read(4)+4;const permutation=[16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15],cl=Array(19).fill(0);for(let i=0;i<nc;i++)cl[permutation[i]]=read(3);const ct=tree(cl),lens=[];while(lens.length<nl+nd){const s=sym(ct);if(s<16)lens.push(s);else{let n,value;if(s===16){if(!lens.length)throw Error('Bad repeat');n=read(2)+3;value=lens[lens.length-1];}else if(s===17){n=read(3)+3;value=0;}else if(s===18){n=read(7)+11;value=0;}else throw Error('Bad length');if(lens.length+n>nl+nd)throw Error('Huffman length overflow');while(n--)lens.push(value);}}if(!lens[256])throw Error('Missing end symbol');lit=tree(lens.slice(0,nl));dist=tree(lens.slice(nl));}
  else throw Error('Reserved block');
  for(;;){const s=sym(lit);if(s<256){put(s);continue;}if(s===256)break;if(s<257||s>285)throw Error('Length symbol');const j=s-257,len=lb[j]+read(le[j]),d=sym(dist);if(d>29)throw Error('Distance symbol');const offset=db[d]+read(de[d]);if(offset>out.length||out.length+len>limit)throw Error('Backreference/output bound');for(let k=0;k<len;k++)put(out[out.length-offset]);}
 }while(!final);return out;
}
function unzipBounded_(blob){
 const b=blob.getBytes().map(x=>x&255);zipBound_(b);const le=(i,n)=>{let v=0;for(let j=0;j<n;j++)v+=b[i+j]*Math.pow(256,j);return v;};let e=-1;for(let i=b.length-22;i>=Math.max(0,b.length-65557);i--)if(le(i,4)===0x06054b50){e=i;break;}let p=le(e+16,4),total=0;const count=le(e+10,2),entries=[];
 function crc(bytes){let c=-1;for(const v of bytes){c^=v;for(let j=0;j<8;j++)c=(c>>>1)^((c&1)?0xedb88320:0);}return (c^-1)>>>0;}
 for(let i=0;i<count;i++){
  const method=le(p+10,2),compressed=le(p+20,4),size=le(p+24,4),nl=le(p+28,2),offset=le(p+42,4),expected=le(p+16,4),name=String.fromCharCode.apply(null,b.slice(p+46,p+46+nl));
  if(le(offset,4)!==0x04034b50||le(offset+8,2)!==method||le(offset+6,2)&1||offset+30>p)throw Error('Invalid local ZIP header');
  const start=offset+30+le(offset+26,2)+le(offset+28,2);if(start+compressed>le(e+16,4)||compressed>5*1024*1024||size>20*1024*1024-total)throw Error('ZIP content bounds');
  const raw=b.slice(start,start+compressed);let out;if(method===0)out=raw;else if(method===8)out=inflateBounded_(raw,size);else throw Error('Unsupported ZIP method');
  if(out.length!==size||crc(out)!==expected)throw Error('ZIP integrity mismatch');total+=out.length;
  if(!name.endsWith('/'))entries.push(Utilities.newBlob(out,'application/xml',name));p+=46+nl+le(p+30,2)+le(p+32,2);
 }
 return entries;
}

function validatePendingControl_(id){
 const f=api_('https://www.googleapis.com/drive/v3/files/'+id+'?fields=id,name,mimeType,parents,owners(emailAddress),permissions(type,role,emailAddress),trashed');
 validateSharing_(f);
 if(f.trashed||f.name!=='Anytime Fitness Reports — Sync & Index'||f.mimeType!=='application/vnd.google-apps.spreadsheet'||!(f.owners||[]).some(o=>o.emailAddress===CFG.account)||!f.parents||f.parents.length!==1)throw Error('Invalid pending control sheet');
 if(f.parents[0]!==CFG.root){
  const home=api_('https://www.googleapis.com/drive/v3/files/root?fields=id');
  if(f.parents[0]!==home.id)throw Error('Pending control outside expected parent');
 }
 validateFolder_(CFG.root);
}

/* Static-only XLS parser. Vendored SheetJS 0.20.3; no conversion or evaluation. */
function staticLegacyWorkbook_(blob){
 const bytes=blob.getBytes().map(b=>b&255);
 if(bytes.length>2*1024*1024)throw Error('Legacy workbook size bound');
 const ole=[208,207,17,224,161,177,26,225].every((b,i)=>bytes[i]===b),prefix=String.fromCharCode.apply(null,bytes.slice(0,100));
 if(!ole&&!/^\s*<\?xml/.test(prefix))throw Error('Unsupported XLS encoding');
 if(ole){
  const u16=i=>bytes[i]+256*bytes[i+1],u32=i=>bytes[i]+256*bytes[i+1]+65536*bytes[i+2]+16777216*bytes[i+3];
  if(bytes.length<512||u16(28)!==65534||![9,12].includes(u16(30))||u16(32)!==6)throw Error('Unsupported OLE geometry');
  const sectors=Math.floor(bytes.length/Math.pow(2,u16(30)));
  if(u32(44)>sectors||u32(64)>sectors||u32(72)>sectors)throw Error('OLE allocation bound');
 }else{
  const xml=blob.getDataAsString();
  if(/<!DOCTYPE|<!ENTITY/i.test(xml)||(xml.match(/<Row\b/g)||[]).length>50000||(xml.match(/<Cell\b/g)||[]).length>100000)throw Error('Unsafe XML or structural bounds');
 }
 const book=XLSX.read(bytes,{type:'array',cellFormula:true,cellHTML:false,cellStyles:false,bookVBA:true,sheetRows:50001});
 if(book.vbaraw||book.SheetNames.length>30)throw Error('Active content or sheet bound');
 let count=0;
 return book.SheetNames.map(name=>{
  const sh=book.Sheets[name],range=XLSX.utils.decode_range(sh['!fullref']||sh['!ref']||'A1');
  if(range.e.r>50000||range.e.c>199)throw Error('Worksheet bounds');
  for(const key of Object.keys(sh))if(key[0]!=='!'){
   const c=sh[key];if(++count>200000||c.f||c.F)throw Error('Formula or cell bounds');
  }
  // Raw cell values are data only. Hyperlinks, comments, macros, and styles are ignored.
  return {name,rows:XLSX.utils.sheet_to_json(sh,{header:1,raw:true,defval:'',blankrows:true}).map(r=>r.map(v=>String(v==null?'':v)))};
 });
}
var CLUBOS_TYPES={'Scheduled Vitals Report':'Vitals','Scheduled Unpaid Invoices Report':'Unpaid Invoices','Scheduled Employee Active Clients Report':'Active Clients','Scheduled Lead Source Report':'Lead Source','Scheduled PT Transactions Report':'PT Transactions','Scheduled Summary Accrual Revenue Report':'Accrual Revenue','Daily Scheduled Appointments Made Report':'Appointments Made','Daily Scheduled Memberships Sold Report':'Memberships Sold','Daily Scheduled New Prospect Report':'New Prospects','Daily Scheduled Package Sales Summary Report':'Package Sales','Scheduled Cancellation Report':'Cancellations'};
function parseClubOsStatic_(sheets){
 const clubs=new Set(),periods=new Set(),families=new Set();let bad=false,periodBad=false;
 function nameClub(v){if(v==='Chantilly, VA')return 'Chantilly';if(v==='Herndon, VA')return 'Herndon';return null;}
 function idClub(v){if(v==='3276')return 'Chantilly';if(v==='5456')return 'Herndon';return null;}
 function check(v,kind){v=abcNorm(v);if(!v||/^(?:Location|Location Name|AF Location ID|Location External Billing ID|Grand Total(?:s(?:\/Averages)?)?|Totals?|Average|All Locations)$/i.test(v))return null;const c=kind==='id'?idClub(v):nameClub(v);if(!c)bad=true;else clubs.add(c);return c;}
 for(const sheet of sheets){
  const rows=sheet.rows.map(r=>r.map(abcNorm)),title=(rows[0]||[])[0];
  if(!CLUBOS_TYPES[title])return abcReview('Unrecognized ClubOS title');families.add(CLUBOS_TYPES[title]);
  if(nameClub(sheet.name))clubs.add(nameClub(sheet.name));else if(/,\s*[A-Z]{2}$/.test(sheet.name))bad=true;
  for(const row of rows.slice(0,4))for(const v of row){
   const m=/^Date Range: (\d{4})-(\d{1,2})-(\d{1,2}) - (\d{4})-(\d{1,2})-(\d{1,2})$/.exec(v);
   if(m){const a=m[1]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[3]).slice(-2),b=m[4]+'-'+('0'+m[5]).slice(-2)+'-'+('0'+m[6]).slice(-2);if(!validDate_(a)||!validDate_(b)||a>b)periodBad=true;else periods.add(a+'|'+b);}
  }
  let names=[],ids=[],horizontalNames=null;
  for(const row of rows){
   // Accrual reports label each section with an explicit paired club name and ID.
   if(title==='Scheduled Summary Accrual Revenue Report'){
    for(const v of row){const m=/^(.+?) \| Ext\. Billing ID:(.+)$/.exec(v);if(m){const a=check(m[1],'name'),b=check(m[2],'id');if(!a||a!==b)bad=true;}}
   }
   if(row[0]==='Personal Training'&&row.slice(1).every(v=>!v))continue;
   // Vitals-style horizontal location and billing-ID header rows.
   if(title==='Scheduled Vitals Report'&&row[0]==='Location Name'){
    if(horizontalNames)bad=true;horizontalNames=row.slice(1);continue;
   }
   if(title==='Scheduled Vitals Report'&&['AF Location ID','External Billing ID*'].includes(row[0])){
    if(!horizontalNames){bad=true;continue;}
    const values=row.slice(1),width=Math.max(values.length,horizontalNames.length);
    for(let c=0;c<width;c++){
     const a=check(horizontalNames[c]||'','name'),b=check(values[c]||'','id');
     if(a!==b)bad=true;
    }
    horizontalNames=null;continue;
   }
   const nc=[],ic=[];row.forEach((v,c)=>{if(v==='Location'||v==='Location Name')nc.push(c);if(v==='AF Location ID'||v==='Location External Billing ID')ic.push(c);});
   if(nc.length||ic.length){names=nc;ids=ic;continue;}
   const a=names.map(c=>check(row[c]||'','name')).filter(Boolean),b=ids.map(c=>check(row[c]||'','id')).filter(Boolean);
   if(a.length===1&&b.length===1&&a[0]!==b[0])bad=true;
  }
  if(horizontalNames)horizontalNames.forEach(v=>check(v,'name'));
 }
 if(bad)return abcReview('Unknown or conflicting ClubOS location evidence');
 if(families.size!==1)return abcReview('Conflicting ClubOS report types');
 if(!clubs.size)return abcReview('No ClubOS location evidence in workbook contents');
 const type=Array.from(families)[0],club=clubs.size===2?'Both Clubs':Array.from(clubs)[0];
 if(periodBad||periods.size!==1||type==='Appointments Made')return {status:'ROUTED_DATE_REVIEW',reason:type==='Appointments Made'?'Filed to verified location; reporting range endpoint semantics require review':periodBad||periods.size>1?'Filed to verified location; conflicting reporting dates require review':'Filed to verified location; reporting date absent from supported content fields',provider:'ClubOS',club,type,periodStart:null,periodEnd:null,reportDate:null,parserVersion:'clubos-static-v3'};
 const period=Array.from(periods)[0].split('|');
 return {status:'READY',reason:'',provider:'ClubOS',club:clubs.size===2?'Both Clubs':Array.from(clubs)[0],type,periodStart:period[0],periodEnd:period[1],reportDate:period[0]===period[1]?period[0]:null,parserVersion:'clubos-static-v3'};
}

/* Exact bytes of previously inspected PDF originals only; no filename inference. */
var VERIFIED_PDF_REPORTS={
  "717b6ed3bf2a35c09c522f4dc19b2e376b2ca3811401ab7f62be2dc214fe6dc7": {
    "status": "READY",
    "reason": "",
    "provider": "ABC Financial",
    "club": "Chantilly",
    "type": "Monthly Billing",
    "periodStart": "2026-09-01",
    "periodEnd": "2026-09-30",
    "reportDate": null,
    "parserVersion": "verified-pdf-sha256-v1",
    "evidenceFileId": "1E38P0AC2Pk16llznH0fpgu9FUjn1OzAW"
  },
  "9890c91fc6e5a3fd8f99c056abbd9ba83dcb9779ce8ce1b080befbdf8721433e": {
    "status": "READY",
    "reason": "",
    "provider": "ABC Financial",
    "club": "Both Clubs",
    "type": "Annual Fee Audit",
    "periodStart": "2026-10-01",
    "periodEnd": "2026-10-31",
    "reportDate": null,
    "parserVersion": "verified-pdf-sha256-v1",
    "evidenceFileId": "1iMtKRQisak1CugIs240T_ivyHSrrGnvE"
  },
  "75e04f4f0b4afffb1fcc430b5bd06c5de58483fe7152c3a761d1c360b8d6ee93": {
    "status": "READY",
    "reason": "",
    "provider": "ABC Financial",
    "club": "Both Clubs",
    "type": "New Accounts",
    "periodStart": "2026-10-01",
    "periodEnd": "2026-10-20",
    "reportDate": null,
    "parserVersion": "verified-pdf-sha256-v1",
    "evidenceFileId": "1m5sdCoe55gsR6d0yNfweshLBAs5wPy9r"
  },
  "d11fdb8346891100ebe07705cdea7f831c38ff89510f098a7e8f256752465522": {
    "status": "READY",
    "reason": "",
    "provider": "ABC Financial",
    "club": "Both Clubs",
    "type": "Renewal Audit",
    "periodStart": "2026-10-21",
    "periodEnd": "2026-10-21",
    "reportDate": "2026-10-21",
    "parserVersion": "verified-pdf-sha256-v1",
    "evidenceFileId": "1LXpdFJhhgNOaXRy8XLXoda0UXiia0DD8"
  },
  "33efb44dae09f145b05086a978acc17b8e26e285a3b65e47855865254a1129d7": {
    "status": "READY",
    "reason": "",
    "provider": "ABC Financial",
    "club": "Herndon",
    "type": "Account Reconciliation",
    "periodStart": "2026-09-01",
    "periodEnd": "2026-09-30",
    "reportDate": null,
    "parserVersion": "verified-pdf-sha256-v1",
    "evidenceFileId": "1_FSB6Nd9ZkYw-tVBWACD6grmjjq4CoPY"
  },
  "20a1ce69fb97aa91aa5054618a95736731e40eba21d58f1fd36f292758976fc0": {
    "status": "READY",
    "reason": "",
    "provider": "ABC Financial",
    "club": "Both Clubs",
    "type": "Delinquency",
    "periodStart": "2026-10-20",
    "periodEnd": "2026-10-20",
    "reportDate": "2026-10-20",
    "parserVersion": "verified-pdf-sha256-v1",
    "evidenceFileId": "1wJJI2DbB5j0NloSpS5bLv2aoEEB9b6Mt"
  },
  "24c7039ff9510e427c96b15e4d68406783a3dfa9575e27e04da50404152e767c": {
    "status": "READY",
    "reason": "",
    "provider": "ABC Financial",
    "club": "Both Clubs",
    "type": "MTD Billing",
    "periodStart": "2026-10-01",
    "periodEnd": "2026-10-20",
    "reportDate": null,
    "parserVersion": "verified-pdf-sha256-v1",
    "evidenceFileId": "1IN1XihHjFA9HnC1qJGJ5WjsFffREQ7BI"
  },
  "f425a795a7d27081e8534b5d5ae0a413182c2bb9db6ef7f46d070e772b887919": {
    "status": "READY",
    "reason": "",
    "provider": "ABC Financial",
    "club": "Chantilly",
    "type": "Account Reconciliation",
    "periodStart": "2026-09-01",
    "periodEnd": "2026-09-30",
    "reportDate": null,
    "parserVersion": "verified-pdf-sha256-v1",
    "evidenceFileId": "1JBXoFGsOqrct0fJIG-Pkzgms5PVAM4Iw"
  },
  "9f35a3d41b036c8be2567c3f332cf37a94b65875754ca0aaecd7695470781982": {
    "status": "READY",
    "reason": "",
    "provider": "ABC Financial",
    "club": "Herndon",
    "type": "Monthly Billing",
    "periodStart": "2026-09-01",
    "periodEnd": "2026-09-30",
    "reportDate": null,
    "parserVersion": "verified-pdf-sha256-v1",
    "evidenceFileId": "1cwKksLzl6pg8AtrxXO_l4f4NsvSTO6NE"
  },
  "4cefaa064cf94ee57a1f6b26d4f43dbcf62f0b7d705578538907a32790bf2184": {
    "status": "READY",
    "reason": "",
    "provider": "ABC Financial",
    "club": "Both Clubs",
    "type": "Account Adjustments",
    "periodStart": "2026-10-01",
    "periodEnd": "2026-10-20",
    "reportDate": null,
    "parserVersion": "verified-pdf-sha256-v1",
    "evidenceFileId": "13YsjkXoAiC2F1spe2CBrT0-SkIsvEkA3"
  },
  "310b53c0586d60d291665c81e4907e748ca837609c3e6ccffd3918a72a059621": {
    "status": "READY",
    "reason": "",
    "provider": "ABC Financial",
    "club": "Both Clubs",
    "type": "End of Term Skip",
    "periodStart": "2026-10-21",
    "periodEnd": "2026-10-21",
    "reportDate": "2026-10-21",
    "parserVersion": "verified-pdf-sha256-v1",
    "evidenceFileId": "1yPrReY4XjGhnJ4Geb89vpZ7ksIfhEpKF"
  },
  "6fec27c77f265176bcc2aec3ed3970cbf40339b009656d18df4e8bcd3330c5ed": {
    "status": "READY",
    "reason": "",
    "provider": "ABC Financial",
    "club": "Both Clubs",
    "type": "Production",
    "periodStart": "2026-10-01",
    "periodEnd": "2026-10-20",
    "reportDate": null,
    "parserVersion": "verified-pdf-sha256-v1",
    "evidenceFileId": "1hp6U3tftEYmTNuF6pzB57-smdwxRVcVs"
  },
  "0cfddcd8737040876ccd5b5cb035c7c6cd8c532f5d85700c7fa0c941de40dbed": {
    "status": "READY",
    "reason": "",
    "provider": "ABC Financial",
    "club": "Both Clubs",
    "type": "Active Members",
    "periodStart": "2026-10-21",
    "periodEnd": "2026-10-21",
    "reportDate": "2026-10-21",
    "parserVersion": "verified-pdf-sha256-v1",
    "evidenceFileId": "1HFPIAuYqTcONHxwXMgUrWkcWc9I4-j5F"
  }
};
function classifyKnownPdf_(sha){const r=VERIFIED_PDF_REPORTS[sha];return r?JSON.parse(JSON.stringify(r)):{status:"NEEDS_REVIEW",reason:"PDF content not in verified original registry; original only"};}

/* Human Sheets checkbox only. API edits do not trigger this handler. */
function enableSheetSync(){
 const props=PropertiesService.getScriptProperties();
 if(props.getProperty('LIVE_TESTS_PASSED')!=='yes'||props.getProperty('IMPORT_APPROVED')!=='yes')throw Error('Filing must be verified and activated first');
 const lock=LockService.getScriptLock();lock.waitLock(10000);
 try{
  assertAccount_();const ss=control_(),id=ss.getId();validateControl_(id);
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='sheetSyncEdit').forEach(t=>ScriptApp.deleteTrigger(t));
  const trigger=ScriptApp.newTrigger('sheetSyncEdit').forSpreadsheet(id).onEdit().create();
  props.setProperty('SHEET_SYNC_TRIGGER_ID',trigger.getUniqueId());
  const sh=ss.getSheetByName('Sync');sh.getRange('B2').insertCheckboxes().setValue(false);
  sh.getRange('B3').setValue('Check B2 to request a sync. Progress and exclusions appear in Runs and Outcomes. If paused, check again to resume.');
  sh.getRange('B1').setValue('Sync now ready');SpreadsheetApp.flush();
 }finally{lock.releaseLock();}
}
function sheetSyncEdit(e){
 const props=PropertiesService.getScriptProperties(),id=props.getProperty('CONTROL_ID');
 if(!e||!e.source||!e.range||String(e.triggerUid)!==props.getProperty('SHEET_SYNC_TRIGGER_ID')||e.source.getId()!==id||e.range.getSheet().getName()!=='Sync'||e.range.getA1Notation()!=='B2'||e.range.getNumRows()!==1||e.range.getNumColumns()!==1||e.value!=='TRUE')return;
 if(props.getProperty('IMPORT_APPROVED')!=='yes'||props.getProperty('LIVE_TESTS_PASSED')!=='yes')return;
 validateControl_(id);const sh=e.source.getSheetByName('Sync');
 sh.getRange('B2').setValue(false);sh.getRange('B1').setValue('Sync requested');SpreadsheetApp.flush();
 try{
  const result=runSync_(null);
  sh.getRange('B1').setValue(result&&result.state==='BUSY'?'Another sync is running; check again after it finishes':result&&result.state==='PAUSED'?'Paused at saved progress; check B2 to continue':'Sync completed; inspect Runs and Outcomes');
 }catch(err){sh.getRange('B1').setValue('Sync stopped; inspect Runs and Outcomes before retrying');}
}
