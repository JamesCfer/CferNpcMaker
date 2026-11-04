/** Generate the importable n8n Map Builder workflow from shared map helpers. */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  PERSPECTIVES, TOP_DOWN_VARIANTS, MOODS, DIALS, SIZE_BOOST, OUTDOOR, ASPECTS,
  checkDials, dialPhrases, cap, sizeBoost, buildPrompt, pickAspect, finalPrompt,
} from '../../shared/scripts/core/maps/prompt.js';
import { PLAN_PROMPT, PLAN_POSITIONS, validatePlan, parsePlan, planText } from '../../shared/scripts/core/maps/plan.js';
import { CHECK_PROMPT, SCALE_PROMPT, LIGHTS_PROMPT, WALLS_PROMPT, mapDefects } from '../../shared/scripts/core/maps/vision.js';
import { imageSize } from '../../shared/scripts/core/maps/image.js';
import { UNIT_SYSTEMS, FEET_PER_METRE, LIGHT_PRESETS, averageScale, sceneScale, gridFit, toFoundryWalls, toFoundryLights, cropBox } from '../../shared/scripts/core/maps/foundry.js';
import { EDIT_QUALITIES, editSettingsCore, convertTraceCore } from './map-actions.mjs';
import { patreonMembersNode, PATREON_CREDENTIAL } from './patreon-credential-migration.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const store = JSON.parse(fs.readFileSync(path.join(root, 'n8n-dnd5e-store-builder-workflow.json'), 'utf8'));
const clone = value => structuredClone(value);
const FINAL_QUALITIES = {
  standard: { label: 'Standard', model: 'google/gemini-3.1-flash-image', resolution: '2K', scaleCalls: 2, cost: 6 },
  economy: { label: 'Economy', model: 'google/gemini-3.1-flash-image', resolution: '1K', scaleCalls: 1, cost: 4 },
  draft: { label: 'Draft', model: 'black-forest-labs/flux.2-klein-4b', resolution: '2K', scaleCalls: 1, cost: 2 },
  high: { label: 'High detail', model: 'bytedance-seed/seedream-5-0-pro', resolution: '2K', scaleCalls: 1, cost: 6 },
};
const WALLS_COST = 1;
const MOCKUP_COSTS = { 1: 1, 2: 2, 4: 3 };
const sharedPrompt = [
  `const PERSPECTIVES = ${JSON.stringify(PERSPECTIVES)};`,
  `const TOP_DOWN_VARIANTS = ${JSON.stringify(TOP_DOWN_VARIANTS)};`,
  `const MOODS = ${JSON.stringify(MOODS)};`,
  `const DIALS = ${JSON.stringify(DIALS)};`,
  `const SIZE_BOOST = ${JSON.stringify(SIZE_BOOST)};`,
  `const OUTDOOR = new RegExp(${JSON.stringify(OUTDOOR.source)}, ${JSON.stringify(OUTDOOR.flags)});`,
  `const ASPECTS = ${JSON.stringify(ASPECTS)};`,
  checkDials.toString(), dialPhrases.toString(), cap.toString(), sizeBoost.toString(),
  buildPrompt.toString(), pickAspect.toString(), finalPrompt.toString(),
].join('\n');
const sharedScale = [
  `const UNIT_SYSTEMS = ${JSON.stringify(UNIT_SYSTEMS)};`,
  `const FEET_PER_METRE = ${JSON.stringify(FEET_PER_METRE)};`,
  averageScale.toString(), sceneScale.toString(),
].join('\n');
const sharedBake = [
  `const UNIT_SYSTEMS = ${JSON.stringify(UNIT_SYSTEMS)};`,
  `const LIGHT_PRESETS = ${JSON.stringify(LIGHT_PRESETS)};`,
  `const SENSE = { none: 0, normal: 20 };`,
  `const DOOR = { none: 0, door: 1 };`,
  gridFit.toString(), toFoundryWalls.toString(), toFoundryLights.toString(),
  convertTraceCore.toString(),
].join('\n');
const source = name => clone(store.nodes.find(node => node.name === name));
const replaceRefs = value => JSON.parse(JSON.stringify(value).replaceAll('dnd5e-store-builder', 'map-builder'));

export function buildMapWorkflow() {
  const nodes = [];
  const connections = {};
  const add = node => { node.id = `map-${String(nodes.length + 1).padStart(2, '0')}`; node.position = [nodes.length * 240, 0]; nodes.push(node); return node; };
  const copied = (name, renamed = name) => { const node = replaceRefs(source(name)); node.name = renamed; return add(node); };
  const code = (name, jsCode) => add({ name, type: 'n8n-nodes-base.code', typeVersion: 2, parameters: { jsCode } });
  const branch = (name, field, value = true) => {
    const node = clone(source('IF: Rate Limit Exceeded?'));
    node.name = name;
    node.parameters.conditions.conditions[0].id = `map-condition-${nodes.length}`;
    node.parameters.conditions.conditions[0].leftValue = `={{ $json.${field} }}`;
    if (!value) node.parameters.conditions.conditions[0].operator.operation = 'false';
    return add(node);
  };
  const link = (from, to, output = 0) => {
    const slots = connections[from] ||= { main: [] };
    slots.main[output] ||= [];
    slots.main[output].push({ node: to, type: 'main', index: 0 });
  };
  const response = (name, status) => {
    const node = copied('Respond to Webhook', name);
    node.parameters.options.responseCode = status;
    return node;
  };
  const http = (name, url, timeout = 180000) => add({
    name, type: 'n8n-nodes-base.httpRequest', typeVersion: 4.2,
    parameters: {
      method: 'POST', url, authentication: 'predefinedCredentialType', nodeCredentialType: 'openRouterApi',
      sendHeaders: true, headerParameters: { parameters: [{ name: 'X-Title', value: 'Celestia Map Builder' }] },
      sendBody: true, specifyBody: 'json', jsonBody: '={{ JSON.stringify($json.body) }}',
      options: { timeout, response: { response: { neverError: true, fullResponse: false } } },
    },
    onError: 'continueRegularOutput', credentials: { openRouterApi: { id: 'gYYHf4NcuzeBjKSk', name: 'OpenRouter account' } },
  });
  const checkBody = `const CHECK_PROMPT = \`${CHECK_PROMPT}\`;\nfunction checkRequest(image) { return { model:'google/gemini-3.7-flash', response_format:{type:'json_object'}, usage:{include:true}, messages:[{role:'user',content:[{type:'text',text:CHECK_PROMPT},{type:'image_url',image_url:{url:'data:'+image.mediaType+';base64,'+image.b64}}]}] }; }`;
  const parseImage = (previous, extra = '') => `const prior = $('${previous}').first().json;\nconst raw = $json || {};\nconst img = raw.data?.[0];\nconst image = img?.b64_json ? {b64:img.b64_json,mediaType:img.media_type||'image/png'} : null;\nconst costUsd = (prior.costUsd||0)+(Number(raw.usage?.cost)||0);\n${extra}\nreturn [{json:{...prior,image,costUsd,...(image?{}:{ok:false,error:'GENERATION_FAILED',message:'The image provider returned no image.',usesCharged:0})}}];`;
  const fallbackCheck = (name, request) => code(name, `const raw=$json||{};\nconst message=String(raw.error?.message??raw.message??raw.error??'');\nconst status=Number(raw.statusCode??raw.status??raw.httpCode??raw.error?.statusCode??raw.error?.status??raw.error?.httpCode??(message.match(/\\b400\\b/)||[])[0]);\nreturn [{json:{...raw,_resolutionFallback:status===400&&/resolution/i.test(message)&&!!$('${request}').first().json.body.resolution}}];`);
  const fallbackBody = (name, request) => code(name, `const {resolution,...body}=$('${request}').first().json.body;\nreturn [{json:{body}}];`);

  copied('Webhook: /webhook/dnd5e-store-builder', 'Webhook: /webhook/map-builder');
  copied('Webhook OPTIONS dnd5e-store-builder', 'Webhook OPTIONS map-builder');
  copied('IF OPTIONS Preflight');
  copied('Respond CORS Preflight');
  copied('Respond CORS Headers');
  copied('Session Validation Node');
  copied('Data Table: Lookup Session');
  copied('Validate Session & Check Expiration');
  copied('IF: Is Authorized?');
  copied('Build Unauthorized Response');
  copied('Respond: 401 Unauthorized');
  code('Map: Prepare Request', `${sharedPrompt}\n${imageSize.toString()}\n${cropBox.toString()}\n${editSettingsCore.toString()}\nconst EDIT_QUALITIES=${JSON.stringify(EDIT_QUALITIES)};\nconst FINAL_QUALITIES=${JSON.stringify(FINAL_QUALITIES)};\nconst UNIT_SYSTEMS=${JSON.stringify(UNIT_SYSTEMS)};\nconst WALLS_COST=${WALLS_COST};\nconst MOCKUP_COSTS=${JSON.stringify(MOCKUP_COSTS)};\nconst body=$json.body||{};\ntry {\n const action=body.action; if(!['plan','mockups','final','bake','edit','retrace','check'].includes(action)) throw new Error('Unknown map action.');\n const count=action==='mockups'?(body.count===undefined?4:body.count):null;\n if(action==='mockups'&&(!Number.isInteger(count)||!Object.prototype.hasOwnProperty.call(MOCKUP_COSTS,count))) throw new Error('Mockup count must be 1, 2, or 4.');\n if(['bake','edit','retrace','check'].includes(action)){\n  const image=body.image; if(!image||typeof image.b64!=='string'||!/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(image.b64)||!image.b64.length||!/^image\\/[a-z0-9.+-]+$/i.test(image.mediaType||'')) throw new Error('A base64 map image is required.');\n  if(action==='check'){ if(body.perspective!=null&&!PERSPECTIVES[body.perspective]) throw new Error('Unknown perspective.'); return [{json:{...$json,_map:{action,image:{b64:image.b64,mediaType:image.mediaType},perspective:body.perspective||'top-down',cost:0}}}]; }\n  const size=imageSize(Buffer.from(image.b64,'base64')); if(!size) throw new Error('Could not read the map image size.');\n  const common={action,image:{b64:image.b64,mediaType:image.mediaType,...size}};\n  if(action==='edit'){\n   const box=body.box; if(!Array.isArray(box)||box.length!==4||!box.every(n=>Number.isFinite(n)&&n>=0&&n<=1)||box[2]-box[0]<0.02||box[3]-box[1]<0.02) throw new Error('Edit box must be within the image and at least 2% wide and tall.');\n   const change=body.change; if(typeof change!=='string'||!change.trim()||change.length>500) throw new Error('A change of up to 500 characters is required.');\n   const quality=body.quality||'standard'; const [x,y,r,b]=cropBox(box,size.width,size.height,0.15); const crop={x,y,width:r-x,height:b-y}; const settings=editSettingsCore(quality,crop,EDIT_QUALITIES);\n   const gmBox={x:Math.trunc(box[0]*size.width),y:Math.trunc(box[1]*size.height),width:Math.trunc(box[2]*size.width)-Math.trunc(box[0]*size.width),height:Math.trunc(box[3]*size.height)-Math.trunc(box[1]*size.height)};\n   const prompt='This is a cut-out piece of a top-down battle map. '+change.trim()+' Keep the edges of the image, the walls, and the floor layout the same so it fits back into the map. Do not draw any grid lines, text, labels, border, or frame. Keep the same art style, lighting, and straight-down camera angle.';\n   return [{json:{...$json,_map:{...common,quality,cost:settings.cost,model:settings.model,resolution:settings.resolution,crop,box:gmBox,prompt,aspect:pickAspect([crop.width,crop.height])}}}];\n  }\n  const squares=body.squares; if(!Array.isArray(squares)||squares.length!==2||!squares.every(n=>Number.isInteger(n)&&n>0&&n<=200)) throw new Error('Squares must be two positive whole numbers up to 200.');\n  if(!UNIT_SYSTEMS[body.units]) throw new Error('Unknown unit system.');\n  const trace=body.trace&&typeof body.trace==='object'?body.trace:{};\n  return [{json:{...$json,_map:{...common,cost:action==='retrace'?1:0,squares,units:body.units,trace:{walls:Array.isArray(trace.walls)?trace.walls.filter(v=>v&&typeof v==='object'&&!Array.isArray(v)):[],lights:Array.isArray(trace.lights)?trace.lights.filter(v=>v&&typeof v==='object'&&!Array.isArray(v)):[]},mood:body.mood,name:body.name}}}];\n }\n const map=body.map; if(!map||typeof map!=='object') throw new Error('A map is required.');\n for(const key of ['description','perspective','genre','style','mood']) if(typeof map[key]!=='string'||!map[key].trim()) throw new Error('Map '+key+' is required.');\n if(!Array.isArray(map.squares)||map.squares.length!==2||!map.squares.every(n=>Number.isInteger(n)&&n>0)) throw new Error('Map squares must be two positive whole numbers.');\n if(map.dials!=null&&(typeof map.dials!=='object'||Array.isArray(map.dials))) throw new Error('Map dials must be an object.');\n if(map.setting!=null&&!['indoor','outdoor'].includes(map.setting)) throw new Error('Unknown map setting.');\n if(map.view!=null&&!TOP_DOWN_VARIANTS[map.view]) throw new Error('Unknown map view.');\n if(map.layout!=null&&typeof map.layout!=='string') throw new Error('Map layout must be text.');\n const cleanMap={...map,...(typeof map.layout==='string'?{layout:map.layout.trim().slice(0,1200)}:{})};\n if(action==='plan') return [{json:{...$json,_map:{action,map:cleanMap,cost:0,startedAt:Date.now()}}}];\n const quality=action==='final'?(body.quality===undefined?'standard':body.quality):null; const finalSettings=action==='final'?FINAL_QUALITIES[quality]:null; let mockup=null;\n if(action==='final'){\n  if(!Object.prototype.hasOwnProperty.call(FINAL_QUALITIES,quality)) throw new Error('Unknown final quality.');\n  mockup=body.mockup; if(!mockup||typeof mockup.b64!=='string'||!/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(mockup.b64)||!mockup.b64.length) throw new Error('A base64 mockup image is required.');\n  if(mockup.mediaType!=null&&!/^image\\/[a-z0-9.+-]+$/i.test(mockup.mediaType)) throw new Error('Invalid mockup media type.');\n  if(body.units!=null&&!['imperial','metric'].includes(body.units)) throw new Error('Unknown unit system.');\n  if(body.square!=null&&!(Number.isFinite(body.square)&&body.square>0)) throw new Error('Square size must be positive.');\n }\n const prompt=action==='mockups'?'Quick rough mockup for layout only. '+buildPrompt(cleanMap):finalPrompt(cleanMap);\n return [{json:{...$json,_map:{action,quality,count,autoPlan:action==='mockups'&&body.autoPlan===true&&!cleanMap.layout,cost:action==='mockups'?MOCKUP_COSTS[count]:finalSettings.cost+(body.walls===true&&map.perspective==='top-down'?WALLS_COST:0),walls:action==='final'&&body.walls===true&&map.perspective==='top-down',resolution:finalSettings?.resolution,scaleCalls:finalSettings?.scaleCalls,map:cleanMap,mockup:mockup?{b64:mockup.b64,mediaType:mockup.mediaType||'image/png'}:null,units:body.units||'imperial',square:body.square,prompt,aspect:pickAspect(map.squares),startedAt:Date.now()}}}];\n} catch(e){return [{json:{ok:false,error:'BAD_REQUEST',message:e.message,usesCharged:0}}];}`);
  branch('IF: Bad Request?', 'ok').parameters.conditions.conditions[0].leftValue = '={{ $json.error === "BAD_REQUEST" }}';
  response('Respond: 400 Bad Request', 400);
  branch('IF: Bake Action?', 'bake').parameters.conditions.conditions[0].leftValue = '={{ $json._map.action === "bake" }}';
  code('Bake: Prepare', `${sharedBake}\nconst prep=$json._map;\nconst {squares,units,trace,image}=prep;\nconst fit=gridFit(image.width,image.height,squares,100);\nconst unit=UNIT_SYSTEMS[units];\nconst {walls,lights,counts}=convertTraceCore(trace,squares,units,UNIT_SYSTEMS,toFoundryWalls,toFoundryLights);\nconst scene={name:typeof prep.name==='string'&&prep.name.trim()?prep.name.trim():'Battle map',width:fit.width,height:fit.height,padding:0,grid:{type:1,size:100,distance:unit.square,units:unit.units},...(prep.mood==='night'?{environment:{darknessLevel:0.7}}:{})};\nreturn [{json:{ok:true,action:'bake',scene,walls,lights,counts,sharpen:fit.sharpen,image},binary:{map:{data:image.b64,mimeType:image.mediaType,fileName:'map.png'}}}];`);
  add({name:'Bake: Resize Image',type:'n8n-nodes-base.editImage',typeVersion:1,parameters:{operation:'resize',dataPropertyName:'map',width:'={{ $json.scene.width }}',height:'={{ $json.scene.height }}',resizeOption:'ignoreAspectRatio',binaryPropertyName:'map'}});
  code('Bake: Respond', `const prior=$('Bake: Prepare').first().json;\nconst binary=$input.first().binary?.map;\nif(!binary) throw new Error('The resized map image is missing.');\nconst bytes=await this.helpers.getBinaryDataBuffer(0,'map');\nreturn [{json:{ok:true,action:'bake',scene:prior.scene,walls:prior.walls,lights:prior.lights,counts:prior.counts,sharpen:prior.sharpen,image:{b64:bytes.toString('base64'),mediaType:binary.mimeType||'image/png'},usesCharged:0}}];`);
  response('Respond: Bake', 200);
  branch('IF: Plan Action?', 'plan').parameters.conditions.conditions[0].leftValue = '={{ $json._map.action === "plan" }}';
  code('Plan: Build Request', `const prep=$json._map;\nconst prompt=${JSON.stringify(PLAN_PROMPT)};\nconst map=prep.map;\nconst fields={description:map.description,perspective:map.perspective,genre:map.genre,setting:map.setting||'indoor',squares:map.squares};\nreturn [{json:{...$json,body:{model:'google/gemini-3.7-flash',response_format:{type:'json_object'},usage:{include:true},messages:[{role:'user',content:prompt+'\\nMap fields: '+JSON.stringify(fields)}]}}}];`);
  http('OpenRouter: Plan Layout', 'https://openrouter.ai/api/v1/chat/completions');
  code('Plan: Parse', `${sharedPrompt}\nconst PLAN_POSITIONS=${JSON.stringify(PLAN_POSITIONS)};\n${validatePlan.toString()}\n${parsePlan.toString()}\n${planText.toString()}\nconst prior=$('Plan: Build Request').first().json;\nconst raw=$json||{};\ntry { const plan=parsePlan(raw.choices?.[0]?.message?.content); const layout=planText(plan,prior._map.map.description); const map={...prior._map.map,layout}; const _map={...prior._map,map,...(prior._map.action==='mockups'?{prompt:'Quick rough mockup for layout only. '+buildPrompt(map)}:{})}; return [{json:{...prior,_map,plan,layout,costUsd:Number(raw.usage?.cost)||0,ok:true}}]; } catch { return [{json:{ok:false,error:'PLAN_FAILED',message:'The layout plan could not be made.',usesCharged:0}}]; }`);
  branch('IF: Plan OK?', 'ok');
  branch('IF: Standalone Plan?', 'plan').parameters.conditions.conditions[0].leftValue = '={{ $json._map.action === "plan" }}';
  code('Plan: Respond', `const {plan,layout,costUsd}=$json; return [{json:{ok:true,plan,layout,costUsd,rawAreas:plan.rawAreas,...(plan.warning?{warning:plan.warning}:{}),usesCharged:0}}];`);
  response('Respond: Plan', 200);
  response('Respond: 502 Plan Failed', 502);
  code('Rate Limit: Prepare Check', `const TIER_LIMITS={'free':3,'no tier':3,'none':3,'':3,'supporter':25,'professional':100,'premium':200,'superuser':999999};\nconst email=($json.sessionEmail||'').toLowerCase().trim();\nconst sessionKey=($json.sessionKey||'').trim();\nconst lookupKey=email||sessionKey;\nif(!lookupKey)throw new Error('No email or session key found.');\nconst sessionTier=($json.sessionTier||'').toLowerCase().trim();\nconst tier=$json.sessionIsSuperuser===true?'superuser':sessionTier;\nconst monthlyLimit=Object.prototype.hasOwnProperty.call(TIER_LIMITS,tier)?TIER_LIMITS[tier]:3;\nconst currentMonth=new Date().toISOString().substring(0,7);\nreturn [{json:{...$json,sessionTier:tier,_rateLimitCheck:{sessionKey:lookupKey+'::'+currentMonth,email:lookupKey,currentMonth,monthlyLimit,tier}}}];`);
  const patreon = patreonMembersNode('Patreon: Members', [0, 0], 'https://www.patreon.com/api/oauth2/v2/campaigns/9988159/members?include=currently_entitled_tiers&fields[member]=patron_status,currently_entitled_amount_cents,email&fields[tier]=title,amount_cents&page[count]=1000', PATREON_CREDENTIAL);
  add(patreon);
  const prepareCheck = nodes.find(n => n.name === 'Rate Limit: Prepare Check');
  prepareCheck.parameters.jsCode = prepareCheck.parameters.jsCode
    .replace('const TIER_LIMITS=', `const __in = $('Map: Prepare Request').first().json;\nconst TIER_LIMITS=`)
    .replaceAll('$json', '__in')
    .replace("const tier=__in.sessionIsSuperuser===true?'superuser':sessionTier;", `let tier=sessionTier;\nif(__in.sessionIsSuperuser===true)tier='superuser';\nelse try {\n  const pages=$('Patreon: Members').all().map(i=>i.json);\n  if(!pages.length||pages.some(p=>!Array.isArray(p.data)))throw new Error('Patreon lookup failed');\n  const tierMap={};\n  let member=null;\n  for(const page of pages){\n    for(const inc of(page.included||[]))if(inc.type==='tier')tierMap[inc.id]=inc.attributes.title;\n    if(!member&&email)member=page.data.find(m=>(m.attributes?.email||'').toLowerCase()===email)||null;\n  }\n  if(member?.attributes?.patron_status==='active_patron'){\n    const names=(member.relationships?.currently_entitled_tiers?.data||[]).map(t=>tierMap[t.id]).filter(Boolean);\n    const paid=names.filter(t=>t.toLowerCase()!=='free').sort((a,b)=>(TIER_LIMITS[b.toLowerCase()]||0)-(TIER_LIMITS[a.toLowerCase()]||0))[0];\n    if(paid)tier=paid.toLowerCase();\n  }\n} catch {} `);
  copied('Get row(s)');
  code('Rate Limit: Validate', `const originalData=$('Rate Limit: Prepare Check').first().json;\nconst rateLimitData=originalData._rateLimitCheck;\nconst cost=originalData._map.cost;\nconst usageData=$input.all().map(i=>i.json).find(r=>r?.session_key);\nconst currentCount=Number(usageData?.npc_count)||0;\nif(cost>0&&currentCount+cost>rateLimitData.monthlyLimit){\n  const resetDate=new Date();resetDate.setMonth(resetDate.getMonth()+1);resetDate.setDate(1);resetDate.setHours(0,0,0,0);\n  return [{json:{ok:false,rateLimitExceeded:true,error:'RATE_LIMIT_EXCEEDED',message:'A map '+(originalData._map.action==='mockups'?'mockups':originalData._map.action==='final'?'final ('+originalData._map.quality+')':originalData._map.action==='edit'?'edit ('+originalData._map.quality+')':'retrace')+' costs '+cost+' uses. You have used '+currentCount+' of your '+rateLimitData.monthlyLimit+' monthly uses.',currentUsage:currentCount,limit:rateLimitData.monthlyLimit,tier:rateLimitData.tier,resetDate:resetDate.toISOString(),upgradeUrl:'https://www.patreon.com/c/CelestiaTools/membership'}}];\n}\nreturn [{json:{...originalData,_rateLimitCheck:{...rateLimitData,currentCount,usageRecordExists:!!usageData,usageRecordId:usageData?.id||null}}}];`);
  copied('IF: Rate Limit Exceeded?');
  branch('IF: Check Action?', 'check').parameters.conditions.conditions[0].leftValue = '={{ $json._map.action === "check" }}';
  code('Check: Build Request', `${checkBody}\nconst prep=$json._map; return [{json:{body:checkRequest(prep.image)}}];`);
  http('OpenRouter: Standalone Check', 'https://openrouter.ai/api/v1/chat/completions');
  code('Check: Parse', `${mapDefects.toString()}\nconst response=$json||{}; const prep=$('Rate Limit: Validate').first().json._map; try { const raw=JSON.parse(String(response.choices?.[0]?.message?.content||'').trim().replace(/^\x60\x60\x60(?:json)?\\s*/,'').replace(/\x60\x60\x60\\s*$/,'')); const defects=mapDefects(raw,{perspective:prep.perspective}); return [{json:{ok:true,raw,defects,costUsd:Number(response.usage?.cost)||0,usesCharged:0}}]; } catch { return [{json:{ok:false,error:'CHECK_FAILED',message:'The map could not be checked.',usesCharged:0}}]; }`);
  branch('IF: Check OK?', 'ok');
  response('Respond: Check', 200);
  branch('IF: Edit Action?', 'edit').parameters.conditions.conditions[0].leftValue = '={{ $json._map.action === "edit" }}';
  code('Edit: Prepare Crop', `const prep=$json._map; return [{json:{crop:prep.crop},binary:{map:{data:prep.image.b64,mimeType:prep.image.mediaType,fileName:'map.png'}}}];`);
  add({ name:'Edit: Crop Image', type:'n8n-nodes-base.editImage', typeVersion:1, parameters:{ operation:'crop',dataPropertyName:'map',width:'={{ $json.crop.width }}',height:'={{ $json.crop.height }}',positionX:'={{ $json.crop.x }}',positionY:'={{ $json.crop.y }}',binaryPropertyName:'map' } });
  code('Edit: Build Request', `const prep=$('Rate Limit: Validate').first().json._map; const bytes=await this.helpers.getBinaryDataBuffer(0,'map'); const binary=$input.first().binary.map; return [{json:{body:{n:1,model:prep.model,prompt:prep.prompt,aspect_ratio:prep.aspect,resolution:prep.resolution,input_references:[{type:'image_url',image_url:{url:'data:'+(binary.mimeType||'image/png')+';base64,'+bytes.toString('base64')}}]}}}];`);
  http('OpenRouter: Edit Image','https://openrouter.ai/api/v1/images',240000);
  fallbackCheck('Edit: Check Resolution Error','Edit: Build Request');
  branch('IF: Edit Resolution Retry?','_resolutionFallback');
  fallbackBody('Edit: Build Without Resolution','Edit: Build Request');
  http('OpenRouter: Edit Image (no resolution)','https://openrouter.ai/api/v1/images',240000);
  code('Edit: Parse Image', `const prep=$('Rate Limit: Validate').first().json._map; const raw=$json||{}; const img=raw.data?.[0]; const patch=img?.b64_json?{b64:img.b64_json,mediaType:img.media_type||'image/png'}:null; return [{json:patch?{ok:true,action:'edit',patch,crop:prep.crop,box:prep.box,costUsd:Number(raw.usage?.cost)||0}:{ok:false,error:'GENERATION_FAILED',message:'The image provider returned no patch.',usesCharged:0}}];`);
  branch('IF: Edit Image OK?','ok');
  branch('IF: Retrace Action?','retrace').parameters.conditions.conditions[0].leftValue = '={{ $json._map.action === "retrace" }}';
  code('Retrace: Build Requests', `const prep=$('Rate Limit: Validate').first().json._map; const WALLS_PROMPT=\`${WALLS_PROMPT}\`; const LIGHTS_PROMPT=\`${LIGHTS_PROMPT}\`; return [['walls',WALLS_PROMPT],['lights',LIGHTS_PROMPT]].map(([purpose,prompt])=>({json:{purpose,body:{model:'google/gemini-3.7-flash',response_format:{type:'json_object'},usage:{include:true},messages:[{role:'user',content:[{type:'text',text:prompt},{type:'image_url',image_url:{url:'data:'+prep.image.mediaType+';base64,'+prep.image.b64}}]}]}}}));`);
  http('OpenRouter: Retrace','https://openrouter.ai/api/v1/chat/completions');
  code('Retrace: Collect', `${sharedBake}\nconst prep=$('Rate Limit: Validate').first().json._map; const requests=$('Retrace: Build Requests').all().map(i=>i.json); const responses=$input.all().map((i,index)=>({purpose:requests[index]?.purpose,raw:i.json})); const trace={walls:[],lights:[]}; for(const {purpose,raw} of responses){try{const value=JSON.parse(String(raw.choices?.[0]?.message?.content||'').trim().replace(/^\x60\x60\x60(?:json)?\\s*/,'').replace(/\x60\x60\x60\\s*$/,''));if(Array.isArray(value?.[purpose]))trace[purpose]=value[purpose].filter(v=>v&&typeof v==='object'&&!Array.isArray(v));}catch{}} const {walls,lights,counts}=convertTraceCore(trace,prep.squares,prep.units,UNIT_SYSTEMS,toFoundryWalls,toFoundryLights); const costUsd=responses.reduce((sum,r)=>sum+(Number(r.raw.usage?.cost)||0),0); return [{json:walls.length||lights.length?{ok:true,action:'retrace',walls,lights,counts,costUsd}:{ok:false,error:'GENERATION_FAILED',message:'No walls or lights could be traced.',usesCharged:0}}];`);
  branch('IF: Retrace OK?','ok');
  branch('Switch: Action', 'action').parameters.conditions.conditions[0].leftValue = '={{ $json._map.action === "mockups" }}';
  branch('IF: Auto Plan?', 'autoPlan').parameters.conditions.conditions[0].leftValue = '={{ $json._map.autoPlan === true }}';
  code('Mockups: Build Requests', `const prep=$json._map?$json:$('Rate Limit: Validate').first().json;\nconst {prompt,aspect,count}=prep._map;\nreturn Array.from({length:count},()=>({json:{body:{n:1,model:'black-forest-labs/flux.2-klein-4b',prompt,aspect_ratio:aspect,resolution:'1K'}}}));`);
  http('OpenRouter: Mockup Images', 'https://openrouter.ai/api/v1/images');
  code('Mockups: Collect', `let planned; try { const item=$('Plan: Parse').first().json; if(item?.ok&&item._map?.action==='mockups') planned=item; } catch {}\nconst prep=planned?._map||$('Rate Limit: Validate').first().json._map;\nconst responses=$input.all().map(i=>i.json);\nconst images=responses.flatMap(r=>(r.data||[]).filter(i=>i?.b64_json).map(i=>({b64:i.b64_json,mediaType:i.media_type||'image/png'})));\nconst costUsd=(planned?.costUsd||0)+responses.reduce((sum,r)=>sum+(Number(r.usage?.cost)||0),0);\nconst failed=prep.count-images.length;\nreturn [{json:images.length?{ok:true,action:'mockups',images,failed,aspect:prep.aspect,prompt:prep.prompt,...(planned?{layout:planned.layout}:{}),costUsd,ms:Date.now()-prep.startedAt}:{ok:false,error:'GENERATION_FAILED',message:'No mockup images were generated.',usesCharged:0}}];`);
  branch('IF: Mockups OK?', 'ok');
  code('Final: Build Request', `const prep=$('Rate Limit: Validate').first().json._map;\nconst {prompt,aspect,mockup}=prep;\nconst quality=${JSON.stringify(FINAL_QUALITIES)}[prep.quality];\nreturn [{json:{body:{n:1,model:quality.model,prompt,aspect_ratio:aspect,resolution:quality.resolution,input_references:[{type:'image_url',image_url:{url:'data:'+mockup.mediaType+';base64,'+mockup.b64}}]},costUsd:0}}];`);
  http('OpenRouter: Final Image', 'https://openrouter.ai/api/v1/images', 240000);
  fallbackCheck('Final: Check Resolution Error', 'Final: Build Request');
  branch('IF: Final Resolution Retry?', '_resolutionFallback');
  fallbackBody('Final: Build Without Resolution', 'Final: Build Request');
  http('OpenRouter: Final Image (no resolution)', 'https://openrouter.ai/api/v1/images', 240000);
  code('Final: Parse Image', parseImage('Final: Build Request'));
  branch('IF: Final Image OK?', 'ok').parameters.conditions.conditions[0].leftValue = '={{ !!$json.image }}';
  code('Final: Build Check', `${checkBody}\nreturn [{json:{body:checkRequest($json.image)}}];`);
  http('OpenRouter: Check', 'https://openrouter.ai/api/v1/chat/completions');
  code('Final: Judge', `${mapDefects.toString()}\nconst prior=$('Final: Parse Image').first().json;\nconst response=$json||{};\nlet defects={retry:false,reasons:[]}, warning;\ntry { const raw=JSON.parse(String(response.choices?.[0]?.message?.content||'').trim().replace(/^\x60\x60\x60(?:json)?\\s*/, '').replace(/\x60\x60\x60\\s*$/, '')); defects=mapDefects(raw,{perspective:$('Rate Limit: Validate').first().json._map.map.perspective}); } catch { warning='The map could not be checked.'; }\nreturn [{json:{...prior,defects,warning,costUsd:prior.costUsd+(Number(response.usage?.cost)||0),retried:false}}];`);
  branch('IF: Retry Final?', 'retry').parameters.conditions.conditions[0].leftValue = '={{ $json.defects.retry }}';
  code('Final: Build Retry', `const prep=$('Rate Limit: Validate').first().json._map;\nconst quality=${JSON.stringify(FINAL_QUALITIES)}[prep.quality];\nreturn [{json:{body:{n:1,model:quality.model,prompt:prep.prompt,aspect_ratio:prep.aspect,resolution:quality.resolution,input_references:[{type:'image_url',image_url:{url:'data:'+prep.mockup.mediaType+';base64,'+prep.mockup.b64}}]}}}];`);
  http('OpenRouter: Final Image (retry)', 'https://openrouter.ai/api/v1/images', 240000);
  fallbackCheck('Final: Check Retry Resolution Error', 'Final: Build Retry');
  branch('IF: Retry Resolution Retry?', '_resolutionFallback');
  fallbackBody('Final: Build Retry Without Resolution', 'Final: Build Retry');
  http('OpenRouter: Final Image (retry, no resolution)', 'https://openrouter.ai/api/v1/images', 240000);
  code('Final: Parse Retry', `const first=$('Final: Judge').first().json;\nconst raw=$json||{};\nconst img=raw.data?.[0];\nconst retryImage=img?.b64_json?{b64:img.b64_json,mediaType:img.media_type||'image/png'}:null;\nreturn [{json:{...first,image:retryImage||first.image,retried:!!retryImage,retryImageFailed:!retryImage,warning:retryImage?first.warning:'The retry image could not be generated; the first image was kept.',costUsd:first.costUsd+(Number(raw.usage?.cost)||0)}}];`);
  code('Final: Build Retry Check', `${checkBody}\nreturn [{json:{body:checkRequest($json.image)}}];`);
  http('OpenRouter: Check (retry)', 'https://openrouter.ai/api/v1/chat/completions');
  code('Final: Judge Retry', `${mapDefects.toString()}\nconst prior=$('Final: Parse Retry').first().json;\nconst response=$json||{};\nlet defects=prior.defects, warning=prior.warning;\nif(!prior.retryImageFailed) { try { const raw=JSON.parse(String(response.choices?.[0]?.message?.content||'').trim().replace(/^\x60\x60\x60(?:json)?\\s*/, '').replace(/\x60\x60\x60\\s*$/, '')); defects=mapDefects(raw,{perspective:$('Rate Limit: Validate').first().json._map.map.perspective}); if(defects.retry) warning='The final map still has: '+defects.reasons.join(', ')+'.'; } catch { warning='The map could not be checked.'; } }\nreturn [{json:{...prior,defects,warning,costUsd:prior.costUsd+(Number(response.usage?.cost)||0)}}];`);
  code('Scale: Build Requests', `const chosen=$json;
const prep=$('Rate Limit: Validate').first().json._map;
const SCALE_PROMPT=\`${SCALE_PROMPT}\`;
const LIGHTS_PROMPT=\`${LIGHTS_PROMPT}\`;
const WALLS_PROMPT=\`${WALLS_PROMPT}\`;
const image=chosen.image;
const tasks=[...Array.from({length:prep.scaleCalls},()=>({purpose:'scale',prompt:SCALE_PROMPT})),...(prep.walls?[{purpose:'lights',prompt:LIGHTS_PROMPT},{purpose:'walls',prompt:WALLS_PROMPT}]:[])];
return tasks.map(({purpose,prompt})=>({json:{purpose,chosen,body:{model:'google/gemini-3.7-flash',response_format:{type:'json_object'},usage:{include:true},messages:[{role:'user',content:[{type:'text',text:prompt},{type:'image_url',image_url:{url:'data:'+image.mediaType+';base64,'+image.b64}}]}]}}}));`);
  http('OpenRouter: Scale', 'https://openrouter.ai/api/v1/chat/completions');
  code('Scale: Combine', `${sharedScale}
const WALLS_COST=${WALLS_COST};
const prep=$('Rate Limit: Validate').first().json._map;
const requests=$('Scale: Build Requests').all().map(i=>i.json);
const chosen=requests[0].chosen;
const responses=$input.all().map((i,index)=>({purpose:requests[index]?.purpose,raw:i.json}));
const parsed=responses.map(({purpose,raw})=>{try{return {purpose,value:JSON.parse(String(raw.choices?.[0]?.message?.content||'').trim().replace(/^\x60\x60\x60(?:json)?\\s*/, '').replace(/\x60\x60\x60\\s*$/, ''))};}catch{return {purpose,value:null};}});
const estimates=parsed.filter(r=>r.purpose==='scale').map(r=>r.value).filter(e=>e&&Number.isFinite(e.width_m)&&e.width_m>0&&Number.isFinite(e.height_m)&&e.height_m>0);
const trace={walls:[],lights:[]};
let measured,scene,warning=chosen.warning;
for(const purpose of ['walls','lights']) if(prep.walls) {
  const value=parsed.find(r=>r.purpose===purpose)?.value;
  if(Array.isArray(value?.[purpose])) trace[purpose]=value[purpose].filter(v=>v&&typeof v==='object'&&!Array.isArray(v));
  else warning=[warning,purpose==='walls'?'Walls could not be traced.':'Lights could not be placed.'].filter(Boolean).join(' ');
}
if(estimates.length){const avg=averageScale(estimates[0],estimates[1]);measured={...avg,estimates};scene=sceneScale(avg,{system:prep.units,square:prep.square});}
else {const size=prep.square||(prep.units==='metric'?1.5:5);const units=prep.units==='metric'?'m':'ft';const squares=prep.map.squares;measured={width_m:null,height_m:null,estimates:[]};scene={squares,width:squares[0]*100,height:squares[1]*100,grid:{size:100,distance:size,units}};warning=[warning,'Map size could not be measured; the requested squares were used.'].filter(Boolean).join(' ');}
const costUsd=chosen.costUsd+responses.reduce((sum,r)=>sum+(Number(r.raw.usage?.cost)||0),0);
const chargedCost=prep.cost-(prep.walls&&trace.walls.length===0&&trace.lights.length===0?WALLS_COST:0);
return [{json:{ok:true,action:'final',quality:prep.quality,image:chosen.image,defects:chosen.defects,retried:chosen.retried,...(warning?{warning}:{}),measured,scene,trace,costUsd,chargedCost,ms:Date.now()-prep.startedAt}}];`);
  branch('IF: Scale OK?', 'ok');
  response('Respond: 502 Generation Failed', 502);
  code('Increment Usage Counter', `const rateLimitData=$('Rate Limit: Validate').first().json._rateLimitCheck;\nlet cost=$('Rate Limit: Validate').first().json._map.cost;\ntry { const final=$('Scale: Combine').first().json; if(final?.ok) cost=final.chargedCost; } catch {}\nconst now=new Date().toISOString();\nreturn [{json:{session_key:rateLimitData.sessionKey,month:rateLimitData.currentMonth,npc_count:rateLimitData.currentCount+cost,last_created:now}}];`);
  copied('Upsert Usage Row');
  code('Map: Build Response', `let payload;\nfor(const name of ['Scale: Combine','Mockups: Collect','Edit: Parse Image','Retrace: Collect']) { try { const candidate=$(name).first().json; if(candidate?.ok){payload=candidate;break;} } catch {} }\nconst cost=payload.chargedCost??$('Rate Limit: Validate').first().json._map.cost;\nconst {chargedCost,...publicPayload}=payload;\nreturn [{json:{...publicPayload,usesCharged:cost}}];`);
  response('Respond to Webhook', 200);

  link('Webhook OPTIONS map-builder','IF OPTIONS Preflight');
  link('IF OPTIONS Preflight','Respond CORS Preflight',0); link('IF OPTIONS Preflight','Respond CORS Headers',1);
  link('Webhook: /webhook/map-builder','Session Validation Node');
  link('Session Validation Node','Data Table: Lookup Session'); link('Data Table: Lookup Session','Validate Session & Check Expiration');
  link('Validate Session & Check Expiration','IF: Is Authorized?');
  link('IF: Is Authorized?','Map: Prepare Request',0); link('IF: Is Authorized?','Build Unauthorized Response',1);
  link('Build Unauthorized Response','Respond: 401 Unauthorized');
  link('Map: Prepare Request','IF: Bad Request?'); link('IF: Bad Request?','Respond: 400 Bad Request',0); link('IF: Bad Request?','IF: Bake Action?',1);
  link('IF: Bake Action?','Bake: Prepare',0); link('IF: Bake Action?','IF: Plan Action?',1);
  link('IF: Plan Action?','Plan: Build Request',0); link('IF: Plan Action?','Patreon: Members',1); link('Patreon: Members','Rate Limit: Prepare Check');
  link('Plan: Build Request','OpenRouter: Plan Layout'); link('OpenRouter: Plan Layout','Plan: Parse'); link('Plan: Parse','IF: Plan OK?');
  link('IF: Plan OK?','IF: Standalone Plan?',0); link('IF: Plan OK?','Respond: 502 Plan Failed',1);
  link('IF: Standalone Plan?','Plan: Respond',0); link('IF: Standalone Plan?','Mockups: Build Requests',1); link('Plan: Respond','Respond: Plan');
  link('Bake: Prepare','Bake: Resize Image'); link('Bake: Resize Image','Bake: Respond'); link('Bake: Respond','Respond: Bake');
  link('Rate Limit: Prepare Check','Get row(s)'); link('Get row(s)','Rate Limit: Validate'); link('Rate Limit: Validate','IF: Rate Limit Exceeded?');
  link('IF: Rate Limit Exceeded?','Build Unauthorized Response',0); link('IF: Rate Limit Exceeded?','IF: Check Action?',1);
  link('IF: Check Action?','Check: Build Request',0); link('IF: Check Action?','IF: Edit Action?',1); link('Check: Build Request','OpenRouter: Standalone Check'); link('OpenRouter: Standalone Check','Check: Parse'); link('Check: Parse','IF: Check OK?'); link('IF: Check OK?','Respond: Check',0); link('IF: Check OK?','Respond: 502 Generation Failed',1);
  link('IF: Edit Action?','Edit: Prepare Crop',0); link('IF: Edit Action?','IF: Retrace Action?',1);
  link('Edit: Prepare Crop','Edit: Crop Image'); link('Edit: Crop Image','Edit: Build Request'); link('Edit: Build Request','OpenRouter: Edit Image'); link('OpenRouter: Edit Image','Edit: Check Resolution Error'); link('Edit: Check Resolution Error','IF: Edit Resolution Retry?');
  link('IF: Edit Resolution Retry?','Edit: Build Without Resolution',0); link('IF: Edit Resolution Retry?','Edit: Parse Image',1); link('Edit: Build Without Resolution','OpenRouter: Edit Image (no resolution)'); link('OpenRouter: Edit Image (no resolution)','Edit: Parse Image'); link('Edit: Parse Image','IF: Edit Image OK?'); link('IF: Edit Image OK?','Increment Usage Counter',0); link('IF: Edit Image OK?','Respond: 502 Generation Failed',1);
  link('IF: Retrace Action?','Retrace: Build Requests',0); link('IF: Retrace Action?','Switch: Action',1); link('Retrace: Build Requests','OpenRouter: Retrace'); link('OpenRouter: Retrace','Retrace: Collect'); link('Retrace: Collect','IF: Retrace OK?'); link('IF: Retrace OK?','Increment Usage Counter',0); link('IF: Retrace OK?','Respond: 502 Generation Failed',1);
  link('Switch: Action','IF: Auto Plan?',0); link('Switch: Action','Final: Build Request',1);
  link('IF: Auto Plan?','Plan: Build Request',0); link('IF: Auto Plan?','Mockups: Build Requests',1);
  link('Mockups: Build Requests','OpenRouter: Mockup Images'); link('OpenRouter: Mockup Images','Mockups: Collect'); link('Mockups: Collect','IF: Mockups OK?');
  link('IF: Mockups OK?','Increment Usage Counter',0); link('IF: Mockups OK?','Respond: 502 Generation Failed',1);
  link('Final: Build Request','OpenRouter: Final Image'); link('OpenRouter: Final Image','Final: Check Resolution Error');
  link('Final: Check Resolution Error','IF: Final Resolution Retry?');
  link('IF: Final Resolution Retry?','Final: Build Without Resolution',0); link('IF: Final Resolution Retry?','Final: Parse Image',1);
  link('Final: Build Without Resolution','OpenRouter: Final Image (no resolution)'); link('OpenRouter: Final Image (no resolution)','Final: Parse Image'); link('Final: Parse Image','IF: Final Image OK?');
  link('IF: Final Image OK?','Final: Build Check',0); link('IF: Final Image OK?','Respond: 502 Generation Failed',1);
  link('Final: Build Check','OpenRouter: Check'); link('OpenRouter: Check','Final: Judge'); link('Final: Judge','IF: Retry Final?');
  link('IF: Retry Final?','Final: Build Retry',0); link('IF: Retry Final?','Scale: Build Requests',1);
  link('Final: Build Retry','OpenRouter: Final Image (retry)'); link('OpenRouter: Final Image (retry)','Final: Check Retry Resolution Error');
  link('Final: Check Retry Resolution Error','IF: Retry Resolution Retry?');
  link('IF: Retry Resolution Retry?','Final: Build Retry Without Resolution',0); link('IF: Retry Resolution Retry?','Final: Parse Retry',1);
  link('Final: Build Retry Without Resolution','OpenRouter: Final Image (retry, no resolution)'); link('OpenRouter: Final Image (retry, no resolution)','Final: Parse Retry');
  link('Final: Parse Retry','Final: Build Retry Check'); link('Final: Build Retry Check','OpenRouter: Check (retry)'); link('OpenRouter: Check (retry)','Final: Judge Retry'); link('Final: Judge Retry','Scale: Build Requests');
  link('Scale: Build Requests','OpenRouter: Scale'); link('OpenRouter: Scale','Scale: Combine'); link('Scale: Combine','IF: Scale OK?');
  link('IF: Scale OK?','Increment Usage Counter',0); link('IF: Scale OK?','Respond: 502 Generation Failed',1);
  link('Increment Usage Counter','Upsert Usage Row'); link('Upsert Usage Row','Map: Build Response'); link('Map: Build Response','Respond to Webhook');
  return { name:'Map Builder', settings:{executionOrder:'v1',executionTimeout:300}, nodes, connections,
    _devNote:'POST /webhook/map-builder accepts free plan and bake, mockups, final, charged edit and retrace. OpenRouter requests use the built-in openRouterApi credential. Active Patreon tier is used when available; the session tier is the fallback.' };
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  fs.writeFileSync(path.join(root, 'n8n-map-builder-workflow.json'), JSON.stringify(buildMapWorkflow(), null, 2) + '\n');
}
