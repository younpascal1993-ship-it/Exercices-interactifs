// Google Apps Script – Moukrim Maths 3e AC
// Google Sheets > Extensions > Apps Script > replace Code.gs > Deploy > New deployment > Web app.
// Execute as: Me; Who has access: Anyone.
// Copy the resulting /exec URL into quiz.html: const SCRIPT_URL="...";
const ID="1ddEUu86xFf_JdSdEUfYkVt0vMxoXcu1k_85DFe-M4PU";
function doPost(e){
  try {
    const data=JSON.parse(e.postData.contents);
    if(data.quiz!=="3e-ac-identites-v1"||!["visit","finish"].includes(data.event))throw Error("Invalid");
    const ss=SpreadsheetApp.openById(ID);
    const sid=String(data.sid||"").slice(0,100);
    const name=String(data.name||"Anonyme").replace(/^[=+@\-]/,"'").slice(0,60);
    const lock=LockService.getScriptLock(); lock.waitLock(10000);
    try{
      if(data.event==="visit"){
        ss.getSheetByName("Visites").appendRow([new Date(),sid,name,"quiz.html"]);
      }else{
        const score=Number(data.score), total=Number(data.total);
        if(!Number.isInteger(score)||total!==24||score<0||score>24)throw Error("Invalid score");
        const answers=Array.isArray(data.answers)?data.answers.slice(0,24).join(","):"";
        ss.getSheetByName("Résultats").appendRow([new Date(),sid,name,score,total,score/total,answers,data.quiz]);
      }
    }finally{lock.releaseLock();}
    return ContentService.createTextOutput("ok");
  }catch(err){return ContentService.createTextOutput("error");}
}
