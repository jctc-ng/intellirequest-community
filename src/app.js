import JSZip from "jszip";
import { renderAsync } from "docx-preview";

const $ = id => document.getElementById(id);
const mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const fields = ["exchange","entity","incorporation","principal","license","reference","date","officer","designation","department","email","approver","subjects"];
let fileHandle = null, selectedFile = null, originalFingerprint = null, generatedBlob = null, changed = false;
const status = text => $("status").textContent = text;
const updateStatus = text => $("fileinfo").textContent = text;
const invalidate = () => { generatedBlob = null; changed = false; $("save").disabled = true; $("download").disabled = true; status("Inputs changed. Regenerate the preview before downloading or saving."); };


// Explicit source-to-target mapping is essential for imported documents without placeholders.
const mappingKeys=["exchange","entity","incorporation","principal","license","reference","date","officer","designation","department","email","approver"];
const mappingLabels={exchange:"Exchange",entity:"Registered entity",incorporation:"Incorporation",principal:"Principal address",license:"Licence",reference:"Reference",date:"Request date",officer:"Officer",designation:"Designation",department:"Department",email:"Official email",approver:"Approving officer"};
const mappingInputs = new Map();
function buildMappingInputs(){
  const holder=$("fieldMapping");holder.replaceChildren();
  for(const key of mappingKeys){
    const row=document.createElement("div");row.className="mapping";
    const label=document.createElement("label");label.textContent=mappingLabels[key];
    const old=document.createElement("input");old.type="text";old.placeholder="Existing text, exact match";old.setAttribute("aria-label","Current "+mappingLabels[key]);old.addEventListener("input",invalidate);
    const target=document.createElement("input");target.type="text";target.readOnly=true;target.tabIndex=-1;target.placeholder="Enter new value in the form above";target.setAttribute("aria-label","New "+mappingLabels[key]);
    const main=$("form").elements.namedItem(key);
    main.addEventListener("input",()=>{target.value=main.value;});
    mappingInputs.set(key,old);row.append(label,old,target);holder.append(row);
  }
}
buildMappingInputs();
function paragraphText(xml){
  return [...xml.getElementsByTagNameNS("*","p")].map(p=>[...p.getElementsByTagNameNS("*","t")].map(t=>t.textContent||"").join("")).join("\n");
}
async function detectCurrentValues(){
  if(!selectedFile){updateStatus("Import a Word document first.");return;}
  try{
    const zip=await JSZip.loadAsync(await selectedFile.arrayBuffer());
    const xml=new DOMParser().parseFromString(await zip.file("word/document.xml").async("string"),"application/xml");
    const text=paragraphText(xml);
    const patterns={
      reference:/\b(?:Our\s+Ref(?:erence)?|Request\s+Ref(?:erence)?|Case\s+Ref(?:erence)?)\s*[:#-]\s*([^\r\n]+)/i,
      date:/\bDate\s*:\s*((?:\d{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]+\s+\d{4})|(?:\d{1,2}[/-]\d{1,2}[/-]\d{4})|(?:\d{4}-\d{2}-\d{2}))/i,
      email:/\b([\w.+-]+@[\w.-]+\.[A-Za-z]{2,})\b/,
      entity:/\b(?:For the attention of|Attention to)\s*:[^\n]*\n\s*([^\n]+)/i
    };
    let found=0;
    for(const [key,pattern] of Object.entries(patterns)){
      const input=mappingInputs.get(key);
      if(!input || input.value.trim())continue;
      const m=text.match(pattern);
      if(m&&m[1]&&m[1].trim().length<180){input.value=m[1].trim();found++;}
    }
    updateStatus("Imported: "+selectedFile.name+". "+found+" existing values suggested. Confirm each value; unfilled original fields require manual mapping.");
    invalidate();
  }catch(err){updateStatus("Unable to inspect document fields: "+err.message);}
}
$("detect").addEventListener("click",detectCurrentValues);
function addRule() {
  const line = document.createElement("div"); line.className = "rule";
  const old = document.createElement("input"); old.placeholder = "Find exact old text"; old.setAttribute("aria-label","Find text");
  const replacement = document.createElement("input"); replacement.placeholder = "Replacement text"; replacement.setAttribute("aria-label","Replacement text");
  const remove = document.createElement("button"); remove.textContent = "Remove"; remove.type = "button";
  old.addEventListener("input", invalidate); replacement.addEventListener("input", invalidate);
  remove.addEventListener("click",()=>{line.remove();invalidate();});
  line.append(old,replacement,remove);$("rules").append(line);
}
addRule();
$("addRule").addEventListener("click",addRule);
$("form").addEventListener("input",invalidate);
$("form").addEventListener("change",invalidate);

async function fingerprint(bytes) {
  const digest = await crypto.subtle.digest("SHA-256",bytes);
  return [...new Uint8Array(digest)].map(n=>n.toString(16).padStart(2,"0")).join("");
}


async function setSelectedFile(file,handle=null){
  if(!file || !file.name.toLowerCase().endsWith(".docx")) throw Error("Choose a .docx Word document.");
  if(file.size>12*1024*1024)throw Error("Maximum supported document size is 12 MB.");
  const buffer=await file.arrayBuffer();
  await JSZip.loadAsync(buffer);
  fileHandle=handle;selectedFile=file;originalFingerprint=await fingerprint(buffer);
  generatedBlob=null;changed=false;$("review").hidden=true;$("generate").disabled=false;
  updateStatus("Selected: "+file.name+" ("+Math.round(file.size/1024)+" KB). "+
    (handle?"Can back up and update the original locally.":"Can preview and download an updated local copy."));
}
$("docx-input").addEventListener("change",async event=>{
  try{await setSelectedFile(event.target.files?.[0]||null);}
  catch(err){updateStatus("Cannot read DOCX: "+err.message);}
});
$("choose").addEventListener("click",async()=>{
  if(typeof window.showOpenFilePicker!=="function"){
    $("docx-input").click();return;
  }
  try{
    const [handle]=await window.showOpenFilePicker({multiple:false,
      types:[{description:"Word DOCX",accept:{[mime]:[".docx"]}}]});
    await setSelectedFile(await handle.getFile(),handle);
  }catch(err){if(err.name!=="AbortError")updateStatus("Cannot open file: "+err.message);}
});

function replaceAcrossNodes(para, find, replacement) {
  if (!find || find===replacement) return 0;
  const nodes=[...para.getElementsByTagNameNS("*","t")];
  const collected=nodes.map(n=>n.textContent||"").join("");
  let matches=0, from=0, idx;
  const spots=[];
  while ((idx=collected.indexOf(find,from))>=0) {
    spots.push(idx); from=idx+find.length;
    if (spots.length>1000) throw Error("Too many matches for: "+find.slice(0,70));
  }
  // Apply from end to start, so character offsets before each match do not change.
  for (let si=spots.length-1;si>=0;si--) {
    const start=spots[si], end=start+find.length;
    let off=0, first=-1,last=-1, a=0,b=0;
    for (let i=0;i<nodes.length;i++) {
      const right=off+(nodes[i].textContent||"").length;
      if(first<0 && start<right){first=i;a=start-off;}
      if(first>=0 && end<=right){last=i;b=end-off;break;}
      off=right;
    }
    if(first<0||last<0)continue;
    if(first===last) {
      const s=nodes[first].textContent||"";
      nodes[first].textContent=s.slice(0,a)+replacement+s.slice(b);
    } else {
      const head=nodes[first].textContent||"",tail=nodes[last].textContent||"";
      nodes[first].textContent=head.slice(0,a)+replacement;
      for(let j=first+1;j<last;j++)nodes[j].textContent="";
      nodes[last].textContent=tail.slice(b);
    }
    nodes[first].setAttribute("xml:space","preserve");
    matches++;
  }
  return matches;
}

function valuesAndRules() {
  const fd=new FormData($("form"));
  const entries=Object.fromEntries(fd.entries());
  const pairs=[];
  for(const key of fields){
    let value=(entries[key]||"").trim();
    if(key==="subjects")value=value.split(/\r?\n/).map(x=>x.trim()).filter(Boolean).join("; ");
    if(value) {
      pairs.push(["{{"+key.toUpperCase()+"}}",value],["[["+key.toUpperCase()+"]]",value]);
    }
  }
  if(entries.date) {
    const date=new Date(entries.date+"T12:00:00");
    if(Number.isNaN(date.getTime()))throw Error("Enter a valid date.");
    const day=date.getDate(), suffix=(day%100>=11&&day%100<=13)?"th":({1:"st",2:"nd",3:"rd"}[day%10]||"th");
    const formatted=day+suffix+" "+date.toLocaleString("en-GB",{month:"long",year:"numeric"});
    pairs.push(["{{DATE_LONG}}",formatted],["[[DATE_LONG]]",formatted]);
  }
  const custom=[...$("rules").querySelectorAll(".rule")].map(row=>({
    old:row.children[0].value,new:row.children[1].value
  })).filter(r=>r.old.trim());
  if(custom.length>75)throw Error("Maximum of 75 Find/Replace rules.");
  for(const rule of custom){
    if(rule.old.length>500||rule.new.length>3000)throw Error("A Find/Replace value is too long.");
    if(rule.old===rule.new)continue;
    if(rule.new.includes(rule.old))throw Error("Replacement cannot contain the search phrase itself: "+rule.old.slice(0,60));
    pairs.push([rule.old,rule.new]);
  }
  for(const key of mappingKeys){
    const old=mappingInputs.get(key).value.trim();
    let value=(entries[key]||"").trim();
    if(key==="date"&&/^\d{4}-\d{2}-\d{2}$/.test(value)){
      const d=new Date(value+"T12:00:00");
      const n=d.getDate(),suffix=(n%100>=11&&n%100<=13)?"th":({1:"st",2:"nd",3:"rd"}[n%10]||"th");
      value=n+suffix+" "+d.toLocaleString("en-GB",{month:"long",year:"numeric"});
    }
    if(old&&value&&old!==value){
      if(value.includes(old))throw Error("New value for "+key+" contains old text, which risks recursive matching.");
      pairs.push([old,value]);
    }
  }
  return pairs.sort((a,b)=>b[0].length-a[0].length);
}

async function generate() {
  if(!selectedFile)return;
  $("review").hidden=false;$("save").disabled=true;$("download").disabled=true;generatedBlob=null;changed=false;
  status("Updating the selected Word document entirely within this browser…");
  try {
    const zip=await JSZip.loadAsync(await selectedFile.arrayBuffer());
    const rules=valuesAndRules();const counts=new Map();
    const names=Object.keys(zip.files).filter(name=>/^word\/(document|header\d+|footer\d+|footnotes|endnotes)\.xml$/.test(name));
    for(const name of names){
      const original=await zip.file(name).async("string");
      const xml=new DOMParser().parseFromString(original,"application/xml");
      if(xml.querySelector("parsererror"))throw Error("Invalid Word XML in "+name);
      const paragraphs=[...xml.getElementsByTagNameNS("*","p")];
      let changedFile=false;
      for(const [old,value] of rules){
        let hits=0;
        for(const para of paragraphs)hits+=replaceAcrossNodes(para,old,value);
        if(hits){changedFile=true;counts.set(old,(counts.get(old)||0)+hits);}
      }
      if(changedFile)zip.file(name,new XMLSerializer().serializeToString(xml));
    }
    generatedBlob=await zip.generateAsync({type:"blob",compression:"DEFLATE",compressionOptions:{level:6}});
    const target=$("preview");target.replaceChildren();
    let renderWarning="";
    try{
      await renderAsync(generatedBlob,target,undefined,{
        className:"docx",inWrapper:true,breakPages:true,ignoreWidth:false,
        ignoreHeight:false,renderHeaders:true,renderFooters:true,
        renderFootnotes:true,renderEndnotes:true,renderComments:false,
        useBase64URL:true
      });
    }catch(renderError){
      target.replaceChildren();
      const xmlText=await zip.file("word/document.xml").async("string");
      const documentXml=new DOMParser().parseFromString(xmlText,"application/xml");
      const pre=document.createElement("pre");
      pre.style.cssText="white-space:pre-wrap;background:white;padding:24px;color:#222";
      pre.textContent=[...documentXml.getElementsByTagNameNS("*","p")]
        .map(p=>[...p.getElementsByTagNameNS("*","t")].map(t=>t.textContent||"").join(""))
        .join("\\n");
      target.appendChild(pre);
      renderWarning="The detailed preview renderer failed; showing a text-only fallback. ";
    }
    const results=[...counts].map(([find,count])=>({find,count}));
    const reminders=[
      "Preview is browser-rendered and can differ from Microsoft Word. Reopen the saved file in Word.",
      "Review names, dates, account identifiers, legal statements, signatures and all annexes.",
      "No institution-specific document or legal authority is supplied by this application."
    ];
    if(!results.length) reminders.unshift("No matching fields or replacement phrases were found. Your document may be unchanged.");
    $("report").textContent="Replacements:\n"+JSON.stringify(results,null,2)+"\n\nReview reminders:\n• "+reminders.join("\n• ");
    changed=results.length>0;$("save").disabled=!changed;$("download").disabled=!changed;
    status(renderWarning+(changed?"Preview generated. Check the changes before updating your original file.":"No edits matched. Add placeholders or Find/Replace rules."));
  }catch(err) {
    generatedBlob=null;$("save").disabled=true;$("download").disabled=true;status("Preview error: "+err.message);
  }
}
$("form").addEventListener("submit",event=>{event.preventDefault();generate();});
$("again").addEventListener("click",generate);

$("download").addEventListener("click",()=>{
  if(!generatedBlob||!changed)return;
  const link=document.createElement("a");
  const url=URL.createObjectURL(generatedBlob);
  link.href=url;
  link.download=(selectedFile.name.replace(/\\.docx$/i,"")||"IntelliRequest")+"-generated.docx";
  document.body.appendChild(link);link.click();link.remove();
  setTimeout(()=>URL.revokeObjectURL(url),15000);
  status("Generated DOCX downloaded. Open it in Word to confirm all details and layout.");
});

$("save").addEventListener("click",async()=>{
  if(!generatedBlob||!changed)return;
  if(!window.showSaveFilePicker){
    // Explicit user action: save a copy on browsers without local write permissions.
    const url=URL.createObjectURL(generatedBlob);
    const link=document.createElement("a");
    link.href=url;link.download=selectedFile.name.replace(/\\.docx$/i,"")+"-updated.docx";
    link.click();setTimeout(()=>URL.revokeObjectURL(url),3000);
    status("Your updated Word copy has been saved using the browser download mechanism. The original remains unchanged.");
    return;
  }
  try{
    if(!fileHandle){
      const destination=await window.showSaveFilePicker({
        suggestedName:selectedFile.name.replace(/\\.docx$/i,"")+"-updated.docx",
        types:[{description:"Word DOCX",accept:{[mime]:[".docx"]}}]
      });
      const writer=await destination.createWritable();
      await writer.write(generatedBlob);await writer.close();
      $("save").disabled=true;
      status("Updated DOCX saved to your chosen location. The original was not modified.");
      return;
    }
    if(!confirm("Create a separate backup and replace the selected Word file with the previewed changes?"))return;
    const backup=await window.showSaveFilePicker({
      suggestedName:selectedFile.name.replace(/\\.docx$/i,"")+"-backup.docx",
      types:[{description:"Word DOCX",accept:{[mime]:[".docx"]}}]
    });
    if(await backup.isSameEntry(fileHandle))throw Error("Backup must be a different file.");
    const latest=await fileHandle.getFile();
    if(await fingerprint(await latest.arrayBuffer())!==originalFingerprint)
      throw Error("Original document changed since import. Open and preview it again.");
    const bw=await backup.createWritable();await bw.write(latest);await bw.close();
    const writer=await fileHandle.createWritable();
    await writer.write(generatedBlob);await writer.close();
    selectedFile=await fileHandle.getFile();
    originalFingerprint=await fingerprint(await selectedFile.arrayBuffer());
    $("save").disabled=true;generatedBlob=null;
    status("Original Word document updated on this device and separate backup created. Open in Word to inspect final formatting.");
  }catch(err){if(err.name!=="AbortError")status("Save failed: "+err.message);}
});

if(!selectedFile)updateStatus("Application ready. Select a DOCX using the native file input or Open with edit permission. Files are processed locally.");
