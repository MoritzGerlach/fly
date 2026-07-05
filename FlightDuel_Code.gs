/***** Flight Duel Backend — Google Apps Script
 *  Deploy as Web App:
 *  Execute as: Me
 *  Who has access: Anyone
 *
 *  First setup is done from the dashboard with a PIN + AirLabs API key.
 *****/

var SHEET_NAME = 'Flights';
var AIRPORT_SHEET = 'Airports';
var HEADERS = [
  'id','owner','flight_iata','flight_date','airline_iata','airline_icao','flight_number','flight_icao',
  'dep_iata','dep_icao','dep_lat','dep_lng','arr_iata','arr_icao','arr_lat','arr_lng',
  'dep_scheduled','dep_estimated','dep_actual','arr_scheduled','arr_estimated','arr_actual',
  'duration_min','dep_delay_min','arr_delay_min','distance_km',
  'aircraft_icao','aircraft_model','manufacturer','registration','status','source','fetched_at','notes','raw_json'
];

function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var cb = cleanCallback_(p.callback || 'callback');
  var out;
  try {
    var action = p.action || 'list';
    if (action === 'setup') out = setup_(p);
    else if (action === 'list') out = { ok:true, flights:listFlights_(), sheetUrl:getBook_().getUrl() };
    else if (action === 'addFromApi') out = addFromApi_(p);
    else if (action === 'addManual') out = addManual_(p);
    else if (action === 'refreshAll') out = refreshAll_(p);
    else if (action === 'delete') out = deleteFlight_(p);
    else out = { ok:false, error:'Unknown action: '+action };
  } catch (err) {
    out = { ok:false, error:String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(cb + '(' + JSON.stringify(out) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function setup_(p) {
  if (!p.pin) throw new Error('Setup-PIN fehlt.');
  var props = PropertiesService.getScriptProperties();
  var existingPin = props.getProperty('ADMIN_PIN');
  if (!existingPin) props.setProperty('ADMIN_PIN', String(p.pin));
  else if (String(existingPin) !== String(p.pin)) throw new Error('Falsche Setup-PIN.');
  if (p.apiKey) props.setProperty('AIRLABS_API_KEY', String(p.apiKey).trim());
  var ss = getBook_();
  ensureFlightsSheet_(ss);
  ensureAirportSheet_(ss);
  return { ok:true, sheetUrl:ss.getUrl(), hasApiKey:!!props.getProperty('AIRLABS_API_KEY') };
}

function addFromApi_(p) {
  var owner = cleanOwner_(p.owner);
  var flight = cleanFlightNo_(p.flight);
  var date = String(p.date || '').slice(0,10);
  if (!flight) throw new Error('Flugnummer fehlt.');
  if (!date) throw new Error('Datum fehlt.');
  var obj = lookupFlight_(flight, date);
  obj.id = makeId_();
  obj.owner = owner;
  obj.flight_date = date;
  obj.fetched_at = new Date().toISOString();
  appendFlight_(obj);
  return { ok:true, flight:publicFlight_(obj) };
}

function addManual_(p) {
  var payload = p.payload || '{}';
  try { payload = decodeURIComponent(payload); } catch(e) {}
  var obj = JSON.parse(payload);
  obj.id = obj.id || makeId_();
  obj.owner = cleanOwner_(obj.owner);
  obj.fetched_at = obj.fetched_at || new Date().toISOString();
  obj.source = obj.source || 'manual';
  obj.flight_iata = cleanFlightNo_(obj.flight_iata || obj.flight || '');
  obj.flight_date = String(obj.flight_date || obj.date || '').slice(0,10);
  addCoordsAndDistance_(obj);
  appendFlight_(obj);
  return { ok:true, flight:publicFlight_(obj) };
}

function refreshAll_(p) {
  var sh = ensureFlightsSheet_(getBook_());
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return { ok:true, updated:0 };
  var header = data[0];
  var idCol = header.indexOf('id');
  var flightCol = header.indexOf('flight_iata');
  var dateCol = header.indexOf('flight_date');
  var ownerCol = header.indexOf('owner');
  var updated = 0;
  for (var r = data.length - 1; r >= 1 && updated < 25; r--) {
    var flight = cleanFlightNo_(data[r][flightCol]);
    var date = asDateString_(data[r][dateCol]);
    if (!flight || !date) continue;
    try {
      var fresh = lookupFlight_(flight, date);
      fresh.id = data[r][idCol] || makeId_();
      fresh.owner = cleanOwner_(data[r][ownerCol]);
      fresh.flight_date = date;
      fresh.fetched_at = new Date().toISOString();
      sh.getRange(r+1, 1, 1, HEADERS.length).setValues([objectToRow_(fresh)]);
      updated++;
      Utilities.sleep(250);
    } catch(e) {}
  }
  return { ok:true, updated:updated, flights:listFlights_() };
}

function deleteFlight_(p) {
  var id = String(p.id || '');
  if (!id) throw new Error('ID fehlt.');
  var sh = ensureFlightsSheet_(getBook_());
  var data = sh.getDataRange().getValues();
  var idCol = data[0].indexOf('id');
  for (var i=1;i<data.length;i++) {
    if (String(data[i][idCol]) === id) { sh.deleteRow(i+1); return { ok:true }; }
  }
  return { ok:false, error:'Flug nicht gefunden.' };
}

function lookupFlight_(flight, date) {
  var errors = [];
  var best = null;
  try {
    var hist = callAirLabs_('v10/historical', { flight_iata:flight });
    best = chooseByDate_(hist.response || hist, date);
    if (best) return normalizeFlight_(best, date, 'airlabs_historical', hist);
  } catch(e) { errors.push('Historical: '+e.message); }
  try {
    var sched = callAirLabs_('v9/schedules', { flight_iata:flight });
    best = chooseByDate_(sched.response || sched, date);
    if (best) return normalizeFlight_(best, date, 'airlabs_schedules', sched);
  } catch(e) { errors.push('Schedules: '+e.message); }
  try {
    var live = callAirLabs_('v9/flight', { flight_iata:flight });
    best = live.response || live;
    if (best && !Array.isArray(best)) return normalizeFlight_(best, date, 'airlabs_flight', live);
    best = chooseByDate_(best, date);
    if (best) return normalizeFlight_(best, date, 'airlabs_flight', live);
  } catch(e) { errors.push('Flight: '+e.message); }
  throw new Error('Keine Flugdaten gefunden. ' + errors.join(' | '));
}

function callAirLabs_(path, params) {
  var key = PropertiesService.getScriptProperties().getProperty('AIRLABS_API_KEY');
  if (!key) throw new Error('AirLabs API-Key ist noch nicht gespeichert.');
  params = params || {};
  params.api_key = key;
  var parts = [];
  for (var k in params) if (params[k] !== undefined && params[k] !== '') parts.push(encodeURIComponent(k)+'='+encodeURIComponent(params[k]));
  var url = 'https://airlabs.co/api/' + path + '?' + parts.join('&');
  var res = UrlFetchApp.fetch(url, { muteHttpExceptions:true, followRedirects:true });
  var code = res.getResponseCode();
  var text = res.getContentText();
  if (code < 200 || code >= 300) throw new Error('HTTP '+code);
  var json = JSON.parse(text);
  if (json.error) throw new Error(json.error.message || JSON.stringify(json.error));
  if (json.response === undefined || json.response === null || (Array.isArray(json.response) && json.response.length === 0)) throw new Error('Leere Antwort');
  return json;
}

function chooseByDate_(response, date) {
  if (!response) return null;
  var arr = Array.isArray(response) ? response : [response];
  if (!arr.length) return null;
  var exact = arr.filter(function(x){ return fieldDate_(x) === date || String(x.dep_time || '').slice(0,10) === date || String(x.dep_time_utc || '').slice(0,10) === date; });
  if (exact.length) return exact[0];
  var target = new Date(date + 'T12:00:00Z').getTime();
  arr.sort(function(a,b){ return Math.abs(dateMillis_(a)-target) - Math.abs(dateMillis_(b)-target); });
  return arr[0];
}

function normalizeFlight_(x, requestedDate, source, raw) {
  var obj = {};
  obj.flight_iata = cleanFlightNo_(x.flight_iata || ((x.airline_iata || '') + (x.flight_number || '')));
  obj.flight_icao = x.flight_icao || '';
  obj.airline_iata = x.airline_iata || (obj.flight_iata.match(/^[A-Z0-9]{2}/) || [''])[0];
  obj.airline_icao = x.airline_icao || '';
  obj.flight_number = x.flight_number || obj.flight_iata.replace(/^[A-Z0-9]{2}/,'');
  obj.flight_date = requestedDate || fieldDate_(x);
  obj.dep_iata = x.dep_iata || '';
  obj.dep_icao = x.dep_icao || '';
  obj.arr_iata = x.arr_iata || '';
  obj.arr_icao = x.arr_icao || '';
  obj.dep_scheduled = x.dep_time || x.dep_time_utc || '';
  obj.dep_estimated = x.dep_estimated || x.dep_estimated_utc || '';
  obj.dep_actual = x.dep_actual || x.dep_actual_utc || '';
  obj.arr_scheduled = x.arr_time || x.arr_time_utc || '';
  obj.arr_estimated = x.arr_estimated || x.arr_estimated_utc || '';
  obj.arr_actual = x.arr_actual || x.arr_actual_utc || '';
  obj.duration_min = num_(x.duration);
  obj.dep_delay_min = num_(x.dep_delayed || x.dep_delay || 0);
  obj.arr_delay_min = num_(x.arr_delayed || x.arr_delay || x.delayed || 0);
  obj.aircraft_icao = x.aircraft_icao || '';
  obj.aircraft_model = x.model || x.aircraft_model || '';
  obj.manufacturer = x.manufacturer || '';
  obj.registration = x.reg_number || x.registration || '';
  obj.status = x.status || '';
  obj.source = source;
  obj.raw_json = JSON.stringify(raw).slice(0, 45000);
  addCoordsAndDistance_(obj);
  return obj;
}

function addCoordsAndDistance_(obj) {
  var dep = getAirport_(obj.dep_iata);
  var arr = getAirport_(obj.arr_iata);
  if (dep) { obj.dep_lat = dep.lat; obj.dep_lng = dep.lng; if (!obj.dep_icao && dep.icao_code) obj.dep_icao = dep.icao_code; }
  if (arr) { obj.arr_lat = arr.lat; obj.arr_lng = arr.lng; if (!obj.arr_icao && arr.icao_code) obj.arr_icao = arr.icao_code; }
  if ((!obj.distance_km || Number(obj.distance_km) === 0) && dep && arr) obj.distance_km = haversineKm_(dep.lat, dep.lng, arr.lat, arr.lng);
}

function getAirport_(iata) {
  iata = String(iata || '').toUpperCase();
  if (!iata) return null;
  var ss = getBook_();
  var sh = ensureAirportSheet_(ss);
  var data = sh.getDataRange().getValues();
  for (var i=1;i<data.length;i++) if (String(data[i][0]).toUpperCase() === iata) return { iata_code:iata, icao_code:data[i][1], name:data[i][2], lat:Number(data[i][3]), lng:Number(data[i][4]) };
  try {
    var json = callAirLabs_('v9/airports', { iata_code:iata, _fields:'iata_code,icao_code,name,lat,lng' });
    var r = Array.isArray(json.response) ? json.response[0] : json.response;
    if (r && r.lat && r.lng) {
      sh.appendRow([iata, r.icao_code || '', r.name || '', Number(r.lat), Number(r.lng), new Date().toISOString()]);
      return { iata_code:iata, icao_code:r.icao_code || '', name:r.name || '', lat:Number(r.lat), lng:Number(r.lng) };
    }
  } catch(e) {}
  return fallbackAirport_(iata);
}

function fallbackAirport_(iata) {
  var A = {
    VIE:[48.1103,16.5697],FRA:[50.0379,8.5622],MUC:[48.3538,11.7861],BER:[52.3667,13.5033],HAM:[53.6304,9.9882],DUS:[51.2895,6.7668],CGN:[50.8659,7.1427],STR:[48.6899,9.2219],SZG:[47.7933,13.0043],ZRH:[47.4582,8.5555],AMS:[52.3105,4.7683],CDG:[49.0097,2.5479],ORY:[48.7233,2.3794],LHR:[51.4700,-0.4543],LGW:[51.1537,-0.1821],DUB:[53.4213,-6.2701],BRU:[50.9014,4.4844],CPH:[55.6180,12.6508],ARN:[59.6498,17.9238],OSL:[60.1976,11.1004],HEL:[60.3172,24.9633],MAD:[40.4983,-3.5676],BCN:[41.2974,2.0833],LIS:[38.7742,-9.1342],FCO:[41.8003,12.2389],MXP:[45.6306,8.7281],ATH:[37.9364,23.9445],IST:[41.2753,28.7519],WAW:[52.1657,20.9671],PRG:[50.1008,14.2600],BUD:[47.4369,19.2556],JFK:[40.6413,-73.7781],EWR:[40.6895,-74.1745],BOS:[42.3656,-71.0096],IAD:[38.9531,-77.4565],ATL:[33.6407,-84.4277],MIA:[25.7959,-80.2870],ORD:[41.9742,-87.9073],DFW:[32.8998,-97.0403],DEN:[39.8561,-104.6737],LAX:[33.9416,-118.4085],SFO:[37.6213,-122.3790],SEA:[47.4502,-122.3088],YYZ:[43.6777,-79.6248],YVR:[49.1967,-123.1815],DXB:[25.2532,55.3657],DOH:[25.2731,51.6081],SIN:[1.3644,103.9915],BKK:[13.6900,100.7501],HKG:[22.3080,113.9185],ICN:[37.4602,126.4407],NRT:[35.7720,140.3929],HND:[35.5494,139.7798],SYD:[-33.9399,151.1753]
  };
  if (!A[iata]) return null;
  return { iata_code:iata, icao_code:'', name:'', lat:A[iata][0], lng:A[iata][1] };
}

function getBook_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SPREADSHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch(e) {}
  }
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) { props.setProperty('SPREADSHEET_ID', active.getId()); return active; }
  var ss = SpreadsheetApp.create('Flight Duel Data');
  props.setProperty('SPREADSHEET_ID', ss.getId());
  return ss;
}

function ensureFlightsSheet_(ss) {
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
  var existing = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),HEADERS.length)).getValues()[0];
  var changed = false;
  for (var i=0;i<HEADERS.length;i++) if (existing[i] !== HEADERS[i]) changed = true;
  if (changed) sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
  sh.setFrozenRows(1);
  return sh;
}

function ensureAirportSheet_(ss) {
  var sh = ss.getSheetByName(AIRPORT_SHEET) || ss.insertSheet(AIRPORT_SHEET);
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,6).setValues([['iata_code','icao_code','name','lat','lng','fetched_at']]);
  return sh;
}

function appendFlight_(obj) {
  var sh = ensureFlightsSheet_(getBook_());
  sh.appendRow(objectToRow_(obj));
}

function listFlights_() {
  var sh = ensureFlightsSheet_(getBook_());
  var data = sh.getDataRange().getValues();
  var out = [];
  for (var r=1;r<data.length;r++) {
    var obj = {};
    for (var c=0;c<HEADERS.length;c++) obj[HEADERS[c]] = serialize_(data[r][c]);
    if (obj.id) out.push(publicFlight_(obj));
  }
  out.sort(function(a,b){ return String(b.flight_date).localeCompare(String(a.flight_date)); });
  return out;
}

function objectToRow_(obj) {
  return HEADERS.map(function(h){ return obj[h] !== undefined ? obj[h] : ''; });
}

function publicFlight_(obj) {
  var o = {};
  for (var i=0;i<HEADERS.length;i++) if (HEADERS[i] !== 'raw_json') o[HEADERS[i]] = obj[HEADERS[i]] !== undefined ? obj[HEADERS[i]] : '';
  return o;
}

function cleanOwner_(v) { return String(v) === 'Farid' ? 'Farid' : 'Moritz'; }
function cleanFlightNo_(v) { return String(v || '').toUpperCase().replace(/[^A-Z0-9]/g,''); }
function cleanCallback_(v) { return String(v || 'callback').replace(/[^A-Za-z0-9_.$]/g,''); }
function makeId_() { return 'fd_' + new Date().getTime().toString(36) + '_' + Math.random().toString(36).slice(2,8); }
function num_(v) { if (v === null || v === undefined || v === '') return 0; var n = Number(v); return isNaN(n) ? 0 : n; }
function serialize_(v) { return Object.prototype.toString.call(v) === '[object Date]' ? Utilities.formatDate(v, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'") : v; }
function asDateString_(v) { if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, 'UTC', 'yyyy-MM-dd'); return String(v || '').slice(0,10); }
function fieldDate_(x) { return String(x.dep_time || x.dep_time_utc || x.arr_time || x.arr_time_utc || '').slice(0,10); }
function dateMillis_(x) { var d = new Date(String(x.dep_time_utc || x.dep_time || x.arr_time_utc || x.arr_time || '1970-01-01').replace(' ', 'T') + 'Z'); var t = d.getTime(); return isNaN(t) ? 0 : t; }
function haversineKm_(lat1,lng1,lat2,lng2) { var R=6371, toRad=function(x){return Number(x)*Math.PI/180}; var p1=toRad(lat1), p2=toRad(lat2), dp=toRad(lat2-lat1), dl=toRad(lng2-lng1); var h=Math.sin(dp/2)*Math.sin(dp/2)+Math.cos(p1)*Math.cos(p2)*Math.sin(dl/2)*Math.sin(dl/2); return Math.round(2*R*Math.asin(Math.sqrt(h))); }
