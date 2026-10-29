import JSZip from "jszip";
import { renderAsync } from "docx-preview";

const $ = id => document.getElementById(id);
const mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const fields = ["exchange","entity","incorporation","principal","license","reference","date","officer","designation","department","email","approver","subjects"];
let fileHandle = null, selectedFile = null, originalFingerprint = null, generatedBlob = null, changed = false;
const status = text => $("status").textContent = text;
const updateStatus = text => $("fileinfo").textContent = text;
const invalidate = () => { generatedBlob = null; $("save").disabled = true; status("Inputs changed. Regenerate the preview before saving."); };

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

$("choose").addEventListener("click", async () => {
  if (typeof window.showOpenFilePicker !== "function") {
    updateStatus("Direct local file updating requires recent Chrome or Edge on HTTPS. Please use one of these browsers.");
    return;
  }
  try {
    const [handle] = await window.showOpenFilePicker({
      multiple:false, types:[{description:"Word DOCX",accept:{[mime]:[".docx"]}}]
    });
    const file = await handle.getFile();
    if (!file.name.toLowerCase().endsWith(".docx")) throw Error("Choose a .docx Word document.");
    if (file.size > 12*1024*1024) throw Error("Maximum supported document is 12 MB.");
    const buffer = await file.arrayBuffer();
    await JSZip.loadAsync(buffer); // Validate ZIP structure before accepting.
    fileHandle=handle; selectedFile=file; originalFingerprint=await fingerprint(buffer);
    generatedBlob=null; $("review").hidden=true; $("generate").disabled=false;
    updateStatus("Selected: "+file.name+" ("+Math.round(file.size/1024)+" KB). Document remains on this device.");
  } catch(err) { if (err.name!=="AbortError") updateStatus("Unable to open file: "+err.message); }
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
  return pairs.sort((a,b)=>b[0].length-a[0].length);
}

async function generate() {
  if(!selectedFile)return;
  $("review").hidden=false;$("save").disabled=true;generatedBlob=null;
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
    await renderAsync(generatedBlob,target,undefined,{
      className:"docx",inWrapper:true,breakPages:true,ignoreWidth:false,
      ignoreHeight:false,renderHeaders:true,renderFooters:true,
      renderFootnotes:true,renderEndnotes:true,renderComments:false,
      useBase64URL:true
    });
    const results=[...counts].map(([find,count])=>({find,count}));
    const reminders=[
      "Preview is browser-rendered and can differ from Microsoft Word. Reopen the saved file in Word.",
      "Review names, dates, account identifiers, legal statements, signatures and all annexes.",
      "No institution-specific document or legal authority is supplied by this application."
    ];
    if(!results.length) reminders.unshift("No matching fields or replacement phrases were found. Your document may be unchanged.");
    $("report").textContent="Replacements:\n"+JSON.stringify(results,null,2)+"\n\nReview reminders:\n• "+reminders.join("\n• ");
    changed=results.length>0;$("save").disabled=!changed;
    status(changed?"Preview generated. Check the changes before updating your original file.":"No edits matched. Add placeholders or Find/Replace rules.");
  }catch(err) {
    generatedBlob=null;$("save").disabled=true;status("Preview error: "+err.message);
  }
}
$("form").addEventListener("submit",event=>{event.preventDefault();generate();});
$("again").addEventListener("click",generate);

$("save").addEventListener("click",async()=>{
  if(!fileHandle||!generatedBlob||!changed)return;
  if(!window.showSaveFilePicker){status("In-place saving requires Chrome or Edge with file system access.");return;}
  if(!confirm("Create a separate backup and replace the selected Word file with the previewed changes?"))return;
  try {
    // Invoke the save picker directly in the user gesture, before awaiting any other work.
    const backup=await window.showSaveFilePicker({
      suggestedName:selectedFile.name.replace(/\.docx$/i,"")+"-backup.docx",
      types:[{description:"Word DOCX",accept:{[mime]:[".docx"]}}]
    });
    if(await backup.isSameEntry(fileHandle))throw Error("Backup must have a different filename.");
    const latest=await fileHandle.getFile();
    if(await fingerprint(await latest.arrayBuffer())!==originalFingerprint)
      throw Error("The original document changed since import. Reopen it and regenerate preview.");
    const bw=await backup.createWritable();
    await bw.write(latest);await bw.close();
    const writable=await fileHandle.createWritable();
    await writable.write(generatedBlob);await writable.close();
    selectedFile=await fileHandle.getFile();
    originalFingerprint=await fingerprint(await selectedFile.arrayBuffer());
    $("save").disabled=true;generatedBlob=null;
    status("Your original Word document was updated on this device and a separate backup was created. Open Word to inspect the final formatting.");
  }catch(err){if(err.name!=="AbortError")status("Save failed: "+err.message);}
});
