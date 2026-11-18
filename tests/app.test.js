"use strict";
/* DOM harness for app control/persistence behavior. This is not a browser renderer. */
const {test}=require('node:test'),assert=require('node:assert/strict'),vm=require('node:vm'),fs=require('node:fs'),P=require('../planner.js');
class Element {
 constructor(tag='div'){this.tag=tag;this.children=[];this.listeners={};this.dataset={};this.textContent='';this.disabled=false;this.checked=false;this.value='';this.files=[];}
 append(...items){this.children.push(...items);}
 replaceChildren(...items){this.children=[...items];this.textContent='';}
 addEventListener(type,fn){this.listeners[type]=fn;}
 setAttribute(){}
 click(){if(!this.disabled)return this.listeners.click?.({target:this});}
}
function boot(storage=new Map(),bootstrap=null){
 const nodes=Object.fromEntries([...fs.readFileSync(require.resolve('../index.html'),'utf8').matchAll(/id="([^"]+)"/g)].map(m=>[m[1],new Element()]));
 if(bootstrap){nodes.bootstrap=new Element('script');nodes.bootstrap.textContent=JSON.stringify(bootstrap);}
 const localStorage={getItem:k=>storage.get(k)||null,setItem:(k,v)=>storage.set(k,v)};
 const document={getElementById:id=>nodes[id]||null,createElement:tag=>new Element(tag)};
 const context={window:{FocusPlanner:P,confirm:()=>true},document,localStorage,Blob,URL:{createObjectURL:()=>"blob:test",revokeObjectURL(){}},setTimeout:fn=>fn(),Date,console};
 vm.runInNewContext(fs.readFileSync(require.resolve('../app.js'),'utf8'),context);
 return {nodes,storage,localStorage,state:()=>JSON.parse(storage.get('focus-planner-v0.3')),import:async raw=>{nodes.import.files=[{size:100,text:async()=>JSON.stringify(raw)}];await nodes.import.listeners.change({target:nodes.import});}};
}
test('app imports, generates, approves, saves and restores actual user decisions',async()=>{
 const app=boot();assert.equal(app.nodes.generate.disabled,true);await app.import(P.demo());assert.equal(app.nodes.generate.disabled,false);assert.equal(app.nodes.weeks.children.flatMap(x=>x.children).length,30);
 const before=app.state().sessions.length;app.nodes.approveAll.click();assert.ok(app.state().sessions.some(s=>s.status==='approved'));app.nodes.generate.click();assert.equal(app.state().sessions.length,before);
 const restored=boot(app.storage);assert.deepEqual(restored.state(),app.state());assert.equal(restored.nodes.saveStatus.textContent,'Restored saved plan');assert.match(restored.nodes.reserved.textContent,/h$/);
});
test('malformed input cannot overwrite saved state, and denied storage remains visible',async()=>{
 const app=boot();await app.import(P.demo());const before=JSON.stringify(app.state());await app.import({schema:'bad'});assert.equal(JSON.stringify(app.state()),before);assert.match(app.nodes.notice.textContent,/Import failed/);
 app.localStorage.setItem=()=>{throw Error('quota')};app.nodes.save.click();assert.match(app.nodes.notice.textContent,/could not save/);assert.equal(app.nodes.export.disabled,false);
});
test('bundled saved plan initializes offline, then existing browser decisions take precedence',()=>{
 const plan=P.createPlan(P.demo());P.generate(plan);plan.sessions[0].status='approved';plan.savedAt=new Date().toISOString();const app=boot(new Map(),plan);assert.equal(app.state().sessions[0].status,'approved');
 app.nodes.approveAll.click();const restored=boot(app.storage,P.createPlan(P.demo()));assert.deepEqual(restored.state(),app.state());
});
test('saved outcome and rejection survive plan download/reimport semantics',async()=>{
 const plan=P.createPlan(P.demo());P.generate(plan);plan.sessions[0].status='completed';plan.sessions[1].status='rejected';plan.savedAt=new Date().toISOString();const app=boot();await app.import(plan);assert.equal(app.state().sessions[0].status,'completed');assert.equal(app.state().sessions[1].status,'rejected');assert.match(app.nodes.insights.textContent,/1 of 1/);
});
