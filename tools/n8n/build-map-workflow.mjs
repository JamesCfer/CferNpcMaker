/** Generate the importable n8n Map Builder workflow from shared map helpers. */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  PERSPECTIVES, TOP_DOWN_VARIANTS, MOODS, DIALS, SIZE_BOOST, OUTDOOR, ASPECTS,
  checkDials, dialPhrases, cap, sizeBoost, buildPrompt, pickAspect, finalPrompt,
} from '../../shared/scripts/core/maps/prompt.js';
import { CHECK_PROMPT, SCALE_PROMPT, mapDefects } from '../../shared/scripts/core/maps/vision.js';
import { UNIT_SYSTEMS, FEET_PER_METRE, averageScale, sceneScale } from '../../shared/scripts/core/maps/foundry.js';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const store = JSON.parse(fs.readFileSync(path.join(root, 'n8n-dnd5e-store-builder-workflow.json'), 'utf8'));
const clone = value => structuredClone(value);
const FINAL_QUALITIES = {
  standard: { resolution: '2K', scaleCalls: 2, cost: 6 },
  economy: { resolution: '1K', scaleCalls: 1, cost: 4 },
};
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
  const http = (name, url) => add({
    name, type: 'n8n-nodes-base.httpRequest', typeVersion: 4.2,
    parameters: {
      method: 'POST', url, authentication: 'predefinedCredentialType', nodeCredentialType: 'openRouterApi',
      sendHeaders: true, headerParameters: { parameters: [{ name: 'X-Title', value: 'Celestia Map Builder' }] },
      sendBody: true, specifyBody: 'json', jsonBody: '={{ JSON.stringify($json.body) }}',
      options: { timeout: 180000, response: { response: { neverError: true, fullResponse: false } } },
    },
    onError: 'continueRegularOutput', credentials: { openRouterApi: { id: 'gYYHf4NcuzeBjKSk', name: 'OpenRouter account' } },
  });
  const checkBody = `const CHECK_PROMPT = \`${CHECK_PROMPT}\`;\nfunction checkRequest(image) { return { model:'google/gemini-3.7-flash', response_format:{type:'json_object'}, usage:{include:true}, messages:[{role:'user',content:[{type:'text',text:CHECK_PROMPT},{type:'image_url',image_url:{url:'data:'+image.mediaType+';base64,'+image.b64}}]}] }; }`;
  const parseImage = (previous, extra = '') => `const prior = $('${previous}').first().json;\nconst raw = $json || {};\nconst img = raw.data?.[0];\nconst image = img?.b64_json ? {b64:img.b64_json,mediaType:img.media_type||'image/png'} : null;\nconst costUsd = (prior.costUsd||0)+(Number(raw.usage?.cost)||0);\n${extra}\nreturn [{json:{...prior,image,costUsd,...(image?{}:{ok:false,error:'GENERATION_FAILED',message:'The image provider returned no image.',usesCharged:0})}}];`;

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
  code('Map: Prepare Request', `${sharedPrompt}\nconst FINAL_QUALITIES = ${JSON.stringify(FINAL_QUALITIES)};\nconst body = $json.body || {};\ntry {\n  const action = body.action;\n  if (!['mockups','final'].includes(action)) throw new Error('Unknown map action.');\n  const map = body.map;\n  if (!map || typeof map !== 'object') throw new Error('A map is required.');\n  for (const key of ['description','perspective','genre','style','mood']) if (typeof map[key] !== 'string' || !map[key].trim()) throw new Error('Map '+key+' is required.');\n  if (!Array.isArray(map.squares) || map.squares.length !== 2 || !map.squares.every(n => Number.isInteger(n) && n > 0)) throw new Error('Map squares must be two positive whole numbers.');\n  if (map.dials != null && (typeof map.dials !== 'object' || Array.isArray(map.dials))) throw new Error('Map dials must be an object.');\n  if (map.setting != null && !['indoor','outdoor'].includes(map.setting)) throw new Error('Unknown map setting.');\n  if (map.view != null && !TOP_DOWN_VARIANTS[map.view]) throw new Error('Unknown map view.');\n  const aspect = pickAspect(map.squares);\n  let mockup = null;\n  const quality = action === 'final' ? (body.quality === undefined ? 'standard' : body.quality) : null;\n  const finalSettings = action === 'final' ? FINAL_QUALITIES[quality] : null;\n  if (action === 'final') {\n    if (!Object.prototype.hasOwnProperty.call(FINAL_QUALITIES, quality)) throw new Error('Unknown final quality.');\n    mockup = body.mockup;\n    if (!mockup || typeof mockup.b64 !== 'string' || !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(mockup.b64) || !mockup.b64.length) throw new Error('A base64 mockup image is required.');\n    if (mockup.mediaType != null && !/^image\\/[a-z0-9.+-]+$/i.test(mockup.mediaType)) throw new Error('Invalid mockup media type.');\n    if (body.units != null && !['imperial','metric'].includes(body.units)) throw new Error('Unknown unit system.');\n    if (body.square != null && !(Number.isFinite(body.square) && body.square > 0)) throw new Error('Square size must be positive.');\n  }\n  const prompt = action === 'mockups' ? 'Quick rough mockup for layout only. ' + buildPrompt(map) : finalPrompt(map);\n  return [{json:{...$json,_map:{action,quality,cost:action==='mockups'?3:finalSettings.cost,resolution:finalSettings?.resolution,scaleCalls:finalSettings?.scaleCalls,map,mockup:mockup?{b64:mockup.b64,mediaType:mockup.mediaType||'image/png'}:null,units:body.units||'imperial',square:body.square,prompt,aspect,startedAt:Date.now()}}}];\n} catch (e) { return [{json:{ok:false,error:'BAD_REQUEST',message:e.message,usesCharged:0}}]; }`);
  branch('IF: Bad Request?', 'ok').parameters.conditions.conditions[0].leftValue = '={{ $json.error === "BAD_REQUEST" }}';
  response('Respond: 400 Bad Request', 400);
  code('Rate Limit: Prepare Check', `const TIER_LIMITS={'free':3,'no tier':3,'none':3,'':3,'supporter':25,'professional':100,'premium':200,'superuser':999999};\nconst email=($json.sessionEmail||'').toLowerCase().trim();\nconst sessionKey=($json.sessionKey||'').trim();\nconst lookupKey=email||sessionKey;\nif(!lookupKey)throw new Error('No email or session key found.');\nconst sessionTier=($json.sessionTier||'').toLowerCase().trim();\nconst tier=$json.sessionIsSuperuser===true?'superuser':sessionTier;\nconst monthlyLimit=Object.prototype.hasOwnProperty.call(TIER_LIMITS,tier)?TIER_LIMITS[tier]:3;\nconst currentMonth=new Date().toISOString().substring(0,7);\nreturn [{json:{...$json,sessionTier:tier,_rateLimitCheck:{sessionKey:lookupKey+'::'+currentMonth,email:lookupKey,currentMonth,monthlyLimit,tier}}}];`);
  copied('Get row(s)');
  code('Rate Limit: Validate', `const originalData=$('Rate Limit: Prepare Check').first().json;\nconst rateLimitData=originalData._rateLimitCheck;\nconst cost=originalData._map.cost;\nconst usageData=$input.all().map(i=>i.json).find(r=>r?.session_key);\nconst currentCount=Number(usageData?.npc_count)||0;\nif(currentCount+cost>rateLimitData.monthlyLimit){\n  const resetDate=new Date();resetDate.setMonth(resetDate.getMonth()+1);resetDate.setDate(1);resetDate.setHours(0,0,0,0);\n  return [{json:{ok:false,rateLimitExceeded:true,error:'RATE_LIMIT_EXCEEDED',message:'A map '+(originalData._map.action==='mockups'?'mockups':'final ('+originalData._map.quality+')')+' costs '+cost+' uses. You have used '+currentCount+' of your '+rateLimitData.monthlyLimit+' monthly uses.',currentUsage:currentCount,limit:rateLimitData.monthlyLimit,tier:rateLimitData.tier,resetDate:resetDate.toISOString(),upgradeUrl:'https://www.patreon.com/c/CelestiaTools/membership'}}];\n}\nreturn [{json:{...originalData,_rateLimitCheck:{...rateLimitData,currentCount,usageRecordExists:!!usageData,usageRecordId:usageData?.id||null}}}];`);
  copied('IF: Rate Limit Exceeded?');
  branch('Switch: Action', 'action').parameters.conditions.conditions[0].leftValue = '={{ $json._map.action === "mockups" }}';
  code('Mockups: Build Requests', `const prep=$('Rate Limit: Validate').first().json;\nconst {prompt,aspect}=prep._map;\nreturn Array.from({length:4},()=>({json:{body:{n:1,model:'black-forest-labs/flux.2-klein-4b',prompt,aspect_ratio:aspect,resolution:'1K'}}}));`);
  http('OpenRouter: Mockup Images', 'https://openrouter.ai/api/v1/images');
  code('Mockups: Collect', `const prep=$('Rate Limit: Validate').first().json._map;\nconst responses=$input.all().map(i=>i.json);\nconst images=responses.flatMap(r=>(r.data||[]).filter(i=>i?.b64_json).map(i=>({b64:i.b64_json,mediaType:i.media_type||'image/png'})));\nconst costUsd=responses.reduce((sum,r)=>sum+(Number(r.usage?.cost)||0),0);\nconst failed=4-images.length;\nreturn [{json:images.length?{ok:true,action:'mockups',images,failed,aspect:prep.aspect,prompt:prep.prompt,costUsd,ms:Date.now()-prep.startedAt}:{ok:false,error:'GENERATION_FAILED',message:'No mockup images were generated.',usesCharged:0}}];`);
  branch('IF: Mockups OK?', 'ok');
  code('Final: Build Request', `const prep=$('Rate Limit: Validate').first().json._map;\nconst {prompt,aspect,mockup}=prep;\nreturn [{json:{body:{n:1,model:'google/gemini-3.1-flash-image',prompt,aspect_ratio:aspect,resolution:prep.resolution,input_references:[{type:'image_url',image_url:{url:'data:'+mockup.mediaType+';base64,'+mockup.b64}}]},costUsd:0}}];`);
  http('OpenRouter: Final Image', 'https://openrouter.ai/api/v1/images');
  code('Final: Parse Image', parseImage('Final: Build Request'));
  branch('IF: Final Image OK?', 'ok').parameters.conditions.conditions[0].leftValue = '={{ !!$json.image }}';
  code('Final: Build Check', `${checkBody}\nreturn [{json:{body:checkRequest($json.image)}}];`);
  http('OpenRouter: Check', 'https://openrouter.ai/api/v1/chat/completions');
  code('Final: Judge', `${mapDefects.toString()}\nconst prior=$('Final: Parse Image').first().json;\nconst response=$json||{};\nlet defects={retry:false,reasons:[]}, warning;\ntry { const raw=JSON.parse(String(response.choices?.[0]?.message?.content||'').trim().replace(/^\x60\x60\x60(?:json)?\\s*/, '').replace(/\x60\x60\x60\\s*$/, '')); defects=mapDefects(raw,{perspective:$('Rate Limit: Validate').first().json._map.map.perspective}); } catch { warning='The map could not be checked.'; }\nreturn [{json:{...prior,defects,warning,costUsd:prior.costUsd+(Number(response.usage?.cost)||0),retried:false}}];`);
  branch('IF: Retry Final?', 'retry').parameters.conditions.conditions[0].leftValue = '={{ $json.defects.retry }}';
  code('Final: Build Retry', `const prep=$('Rate Limit: Validate').first().json._map;\nconst first=$('Final: Judge').first().json;\nreturn [{json:{body:{n:1,model:'google/gemini-3.1-flash-image',prompt:prep.prompt,aspect_ratio:prep.aspect,resolution:prep.resolution,input_references:[{type:'image_url',image_url:{url:'data:'+prep.mockup.mediaType+';base64,'+prep.mockup.b64}}]}}}];`);
  http('OpenRouter: Final Image (retry)', 'https://openrouter.ai/api/v1/images');
  code('Final: Parse Retry', `const first=$('Final: Judge').first().json;\nconst raw=$json||{};\nconst img=raw.data?.[0];\nconst retryImage=img?.b64_json?{b64:img.b64_json,mediaType:img.media_type||'image/png'}:null;\nreturn [{json:{...first,image:retryImage||first.image,retried:!!retryImage,retryImageFailed:!retryImage,warning:retryImage?first.warning:'The retry image could not be generated; the first image was kept.',costUsd:first.costUsd+(Number(raw.usage?.cost)||0)}}];`);
  code('Final: Build Retry Check', `${checkBody}\nreturn [{json:{body:checkRequest($json.image)}}];`);
  http('OpenRouter: Check (retry)', 'https://openrouter.ai/api/v1/chat/completions');
  code('Final: Judge Retry', `${mapDefects.toString()}\nconst prior=$('Final: Parse Retry').first().json;\nconst response=$json||{};\nlet defects=prior.defects, warning=prior.warning;\nif(!prior.retryImageFailed) { try { const raw=JSON.parse(String(response.choices?.[0]?.message?.content||'').trim().replace(/^\x60\x60\x60(?:json)?\\s*/, '').replace(/\x60\x60\x60\\s*$/, '')); defects=mapDefects(raw,{perspective:$('Rate Limit: Validate').first().json._map.map.perspective}); if(defects.retry) warning='The final map still has: '+defects.reasons.join(', ')+'.'; } catch { warning='The map could not be checked.'; } }\nreturn [{json:{...prior,defects,warning,costUsd:prior.costUsd+(Number(response.usage?.cost)||0)}}];`);
  code('Scale: Build Requests', `const chosen=$json;\nconst prep=$('Rate Limit: Validate').first().json._map;\nconst SCALE_PROMPT=\`${SCALE_PROMPT}\`;\nconst image=chosen.image;\nconst body={model:'google/gemini-3.7-flash',response_format:{type:'json_object'},usage:{include:true},messages:[{role:'user',content:[{type:'text',text:SCALE_PROMPT},{type:'image_url',image_url:{url:'data:'+image.mediaType+';base64,'+image.b64}}]}]};\nreturn Array.from({length:prep.scaleCalls},()=>({json:{body,chosen}}));`);
  http('OpenRouter: Scale', 'https://openrouter.ai/api/v1/chat/completions');
  code('Scale: Combine', `${sharedScale}\nconst prep=$('Rate Limit: Validate').first().json._map;\nconst chosen=$('Scale: Build Requests').first().json.chosen;\nconst responses=$input.all().map(i=>i.json);\nconst estimates=responses.map(r=>{try{return JSON.parse(String(r.choices?.[0]?.message?.content||'').trim().replace(/^\x60\x60\x60(?:json)?\\s*/, '').replace(/\x60\x60\x60\\s*$/, ''));}catch{return null;}}).filter(e=>e&&Number.isFinite(e.width_m)&&e.width_m>0&&Number.isFinite(e.height_m)&&e.height_m>0);\nlet measured,scene,warning=chosen.warning;\nif(estimates.length){ const avg=averageScale(estimates[0],estimates[1]); measured={...avg,estimates}; scene=sceneScale(avg,{system:prep.units,square:prep.square}); }\nelse { const size=prep.square||(prep.units==='metric'?1.5:5); const units=prep.units==='metric'?'m':'ft'; const squares=prep.map.squares; measured={width_m:null,height_m:null,estimates:[]}; scene={squares,width:squares[0]*100,height:squares[1]*100,grid:{size:100,distance:size,units}}; warning=[warning,'Map size could not be measured; the requested squares were used.'].filter(Boolean).join(' '); }\nconst costUsd=chosen.costUsd+responses.reduce((sum,r)=>sum+(Number(r.usage?.cost)||0),0);\nreturn [{json:{ok:true,action:'final',quality:prep.quality,image:chosen.image,defects:chosen.defects,retried:chosen.retried,...(warning?{warning}:{}),measured,scene,costUsd,ms:Date.now()-prep.startedAt}}];`);
  branch('IF: Scale OK?', 'ok');
  response('Respond: 502 Generation Failed', 502);
  code('Increment Usage Counter', `const rateLimitData=$('Rate Limit: Validate').first().json._rateLimitCheck;\nconst cost=$('Rate Limit: Validate').first().json._map.cost;\nconst now=new Date().toISOString();\nreturn [{json:{session_key:rateLimitData.sessionKey,month:rateLimitData.currentMonth,npc_count:rateLimitData.currentCount+cost,last_created:now}}];`);
  copied('Upsert Usage Row');
  code('Map: Build Response', `let payload;\ntry { payload=$('Scale: Combine').first().json; } catch {}\nif(!payload?.ok) { try { payload=$('Mockups: Collect').first().json; } catch {} }\nconst cost=$('Rate Limit: Validate').first().json._map.cost;\nreturn [{json:{...payload,usesCharged:cost}}];`);
  response('Respond to Webhook', 200);

  link('Webhook OPTIONS map-builder','IF OPTIONS Preflight');
  link('IF OPTIONS Preflight','Respond CORS Preflight',0); link('IF OPTIONS Preflight','Respond CORS Headers',1);
  link('Webhook: /webhook/map-builder','Session Validation Node');
  link('Session Validation Node','Data Table: Lookup Session'); link('Data Table: Lookup Session','Validate Session & Check Expiration');
  link('Validate Session & Check Expiration','IF: Is Authorized?');
  link('IF: Is Authorized?','Map: Prepare Request',0); link('IF: Is Authorized?','Build Unauthorized Response',1);
  link('Build Unauthorized Response','Respond: 401 Unauthorized');
  link('Map: Prepare Request','IF: Bad Request?'); link('IF: Bad Request?','Respond: 400 Bad Request',0); link('IF: Bad Request?','Rate Limit: Prepare Check',1);
  link('Rate Limit: Prepare Check','Get row(s)'); link('Get row(s)','Rate Limit: Validate'); link('Rate Limit: Validate','IF: Rate Limit Exceeded?');
  link('IF: Rate Limit Exceeded?','Build Unauthorized Response',0); link('IF: Rate Limit Exceeded?','Switch: Action',1);
  link('Switch: Action','Mockups: Build Requests',0); link('Switch: Action','Final: Build Request',1);
  link('Mockups: Build Requests','OpenRouter: Mockup Images'); link('OpenRouter: Mockup Images','Mockups: Collect'); link('Mockups: Collect','IF: Mockups OK?');
  link('IF: Mockups OK?','Increment Usage Counter',0); link('IF: Mockups OK?','Respond: 502 Generation Failed',1);
  link('Final: Build Request','OpenRouter: Final Image'); link('OpenRouter: Final Image','Final: Parse Image'); link('Final: Parse Image','IF: Final Image OK?');
  link('IF: Final Image OK?','Final: Build Check',0); link('IF: Final Image OK?','Respond: 502 Generation Failed',1);
  link('Final: Build Check','OpenRouter: Check'); link('OpenRouter: Check','Final: Judge'); link('Final: Judge','IF: Retry Final?');
  link('IF: Retry Final?','Final: Build Retry',0); link('IF: Retry Final?','Scale: Build Requests',1);
  link('Final: Build Retry','OpenRouter: Final Image (retry)'); link('OpenRouter: Final Image (retry)','Final: Parse Retry');
  link('Final: Parse Retry','Final: Build Retry Check'); link('Final: Build Retry Check','OpenRouter: Check (retry)'); link('OpenRouter: Check (retry)','Final: Judge Retry'); link('Final: Judge Retry','Scale: Build Requests');
  link('Scale: Build Requests','OpenRouter: Scale'); link('OpenRouter: Scale','Scale: Combine'); link('Scale: Combine','IF: Scale OK?');
  link('IF: Scale OK?','Increment Usage Counter',0); link('IF: Scale OK?','Respond: 502 Generation Failed',1);
  link('Increment Usage Counter','Upsert Usage Row'); link('Upsert Usage Row','Map: Build Response'); link('Map: Build Response','Respond to Webhook');
  return { name:'Map Builder', settings:{executionOrder:'v1'}, nodes, connections,
    _devNote:'POST /webhook/map-builder accepts {action:"mockups",map} or {action:"final",map,mockup:{b64,mediaType},quality?,units?,square?}. Success returns images or a final image, scene, measured size, costUsd, ms, and usesCharged. Errors return ok:false and charge zero. OpenRouter requests use the built-in openRouterApi credential OpenRouter account (id gYYHf4NcuzeBjKSk). Patreon live tier lookup was intentionally left out; session tier is authoritative.' };
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  fs.writeFileSync(path.join(root, 'n8n-map-builder-workflow.json'), JSON.stringify(buildMapWorkflow(), null, 2) + '\n');
}
