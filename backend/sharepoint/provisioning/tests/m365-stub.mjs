#!/usr/bin/env node
// A stateful CLI contract double. No network, credential store, or tenant calls.
import { readFileSync, writeFileSync } from 'node:fs';
const args = process.argv.slice(2);
const state = JSON.parse(readFileSync(process.env.QSTAR_TEST_STATE, 'utf8'));
const command = args[0] === 'request' ? 'request' : args.slice(0, 3).join(' ');
const value = name => args[args.indexOf(name) + 1];
const options = {};
for (let i = args[0] === 'request' ? 1 : 3; i < args.length; i++) if (args[i].startsWith('--')) {
  options[args[i].slice(2)] = args[i + 1]?.startsWith('--') || i === args.length - 1 ? true : args[++i];
}
const fail = message => { console.error(message); process.exit(1); };
const decode = text => text.replaceAll('&amp;', '&').replaceAll('&lt;', '<').replaceAll('&gt;', '>');
function field(xml) {
  const attributes = Object.fromEntries([...xml.matchAll(/(\w+)=(?:'([^']*)'|"([^"]*)")/g)].map(m => [m[1], m[2] ?? m[3]]));
  return { InternalName: attributes.Name, TypeAsString: attributes.Type, SchemaXml: xml,
    Required: attributes.Required === 'TRUE', Indexed: attributes.Indexed === 'TRUE',
    Choices: [...xml.matchAll(/<CHOICE>(.*?)<\/CHOICE>/g)].map(m => decode(m[1])) };
}
let result;
const title = options.listTitle || options.title;
const list = state.lists[title];
switch (command) {
  case 'request': {
    const parsed = new URL(options.url);
    const path = decodeURIComponent(parsed.pathname + parsed.search).split('/_api/')[1];
    const match = path.match(/getbytitle\('([^']+)'\)/);
    const target = match ? state.lists[match[1]] : undefined;
    if (match && !target) fail('List does not exist');
    const root = title => `/sites/qstar/Lists/${title.replaceAll(' ', '')}`;
    if (path.includes('RootFolder/ServerRelativeUrl')) result = {RootFolder:{ServerRelativeUrl:root(match[1])}};
    else if (path.includes('TaskOwnerId')) result = {value:target.items};
    else if (path.includes('FSObjType')) result = {value:target.items};
    else if (path.includes('currentuser')) result = {Id:99};
    else if (options['x-http-method'] === 'MERGE') Object.assign(target,JSON.parse(options.body));
    else if (path.includes('RootFolder/Folders')) result = {value:(target.folders || []).map(ServerRelativeUrl=>({ServerRelativeUrl}))};
    else if (path.startsWith('web/folders/addUsingPath')) {
      const url = path.match(/DecodedUrl='([^']+)'/)[1];
      const pair = Object.entries(state.lists).find(([name])=>url.startsWith(root(name)+'/'));
      if (!pair) fail('Unexpected folder path');
      pair[1].folders ||= []; if (!pair[1].folders.includes(url)) pair[1].folders.push(url);
    } else fail(`Unexpected request: ${path}`);
    break;
  }
  case 'spo list get': if (!list) fail('List does not exist'); result = { Title: title }; break;
  case 'spo list add':
    state.lists[title] = { fields: { Author: field("<Field Type='User' Name='Author'/>"), Title: field("<Field Type='Text' Name='Title' Required='TRUE' />") }, items: [] };
    break;
  case 'spo field get':
    result = list?.fields[options.internalName];
    if (!result) fail('Field does not exist');
    if (options.query === 'TypeAsString') result = result.TypeAsString;
    break;
  case 'spo field add': {
    if (!list) fail('List does not exist');
    const created = field(options.xml);
    if (state.failField === created.InternalName) fail('403 permission denied');
    const allowed = ['AddFieldInternalNameHint', 'AddFieldToDefaultView'];
    if (options.options.split(',').some(o => !allowed.includes(o))) fail('Invalid field creation option');
    if (list.fields[created.InternalName]) fail('Already exists');
    list.fields[created.InternalName] = created;
    break;
  }
  case 'spo field set': {
    const target = list?.fields[options.internalName];
    if (!target) fail('Field does not exist');
    if (options.SchemaXml) Object.assign(target, field(options.SchemaXml));
    if ('Required' in options) target.Required = options.Required === 'true';
    if ('Indexed' in options) target.Indexed = options.Indexed === 'true';
    break;
  }
  case 'spo listitem list':
    if (options.pageSize || options.pageNumber) fail('Test requires complete automatic paging');
    if (!list) fail('List does not exist');
    result = list.items;
    break;
  case 'spo listitem set': {
    const item = list.items.find(row => row.Id === Number(options.id));
    if (!item) fail('Item does not exist');
    for (const key of ['Region', 'ReferenceOffset']) if (key in options) item[key] = key === 'ReferenceOffset' ? Number(options[key]) : options[key];
    break;
  }
  case 'spo listitem add':
    list.items.push({Id: 1, Title: options.Title, ReferenceOffset: Number(options.ReferenceOffset)});
    break;
  default: fail(`Unexpected command: ${args.join(' ')}`);
}
state.commands.push(args);
writeFileSync(process.env.QSTAR_TEST_STATE, JSON.stringify(state));
if (result !== undefined) console.log(options.output === 'text' && typeof result === 'string' ? result : JSON.stringify(result));
