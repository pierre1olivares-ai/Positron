import assert from "node:assert/strict";
import test from "node:test";

// Real web-part render logic, with only the SPFx host and boundary services substituted.
const Module = require("node:module");
const load = Module._load;
const elements: any[] = [];
const creations: { kind: string; args: any[]; instance: object }[] = [];
function service(kind: string) {
  return class {
    constructor(...args: any[]) { creations.push({ kind, args, instance: this }); }
    async run() { return []; }
  };
}
const dependencies: Record<string, any> = {
  "react-dom": { render: (element: any) => elements.push(element), unmountComponentAtNode: () => undefined },
  "@microsoft/sp-core-library": { Version: { parse: (value: string) => value } },
  "@microsoft/sp-property-pane": { PropertyPaneTextField: () => ({}), PropertyPaneToggle: () => ({}), PropertyPaneChoiceGroup: () => ({}) },
  "@microsoft/sp-webpart-base": { BaseClientSideWebPart: class {} },
  "@microsoft/sp-component-base": {}, QstarIssueManagerWebPartStrings: {},
  "./components/QstarIssueManager": { default: () => null, __esModule: true },
  "./services/SharePointDataService": { SharePointDataService: service("sharepoint") },
  "./services/BackendApiDataService": { BackendApiDataService: service("backend") },
  "./services/BackendApiClient": { BackendApiClient: service("client") },
  "./services/BackendDiagnosticsService": { BackendDiagnosticsService: service("backend-diagnostics") },
  "./services/ConnectionDiagnosticsService": { ConnectionDiagnosticsService: service("direct-diagnostics") },
  "./services/BackendRoleResolver": { BackendRoleResolver: service("backend-role") },
  "./services/SharePointRoleResolver": { SharePointRoleResolver: service("direct-role"), DevelopmentRoleResolver: service("local-role") },
  "./services/MockDataService": { MockDataService: service("mock") },
};
Module._load = function (name: string, ...args: any[]) { return dependencies[name] || load.call(this, name, ...args); };
const WebPart = require("../src/webparts/qstarIssueManager/QstarIssueManagerWebPart").default;
Module._load = load;
const current = () => elements[elements.length - 1];
function webpart(local = false) {
  const part = new WebPart();
  part.context = { isServedFromLocalhost: local, pageContext: { web: { absoluteUrl: "https://tenant.sharepoint.com/sites/direct" }, user: { email: "caller@example.com", displayName: "Caller" } }, sdks: {} };
  part.properties = { description: "", siteUrl: "", issuesListName: "", progressListName: "", betaAccessMode: false };
  part.domElement = {};
  return part;
}

test("default direct mode caches services while target/mode/resource/identity changes remount and refresh role source", () => {
  const part = webpart(); part.render();
  const initial = current();
  assert.equal(initial.props.connection.dataSourceMode, "sharepoint"); assert.equal(initial.props.connection.siteUrl, "https://tenant.sharepoint.com/sites/direct");
  assert.equal(creations.filter(x => x.instance === initial.props.dataService)[0].kind, "sharepoint");
  const count = creations.length;
  part.properties.description = "Unrelated edit"; part.render();
  assert.equal(current().key, initial.key); assert.equal(current().props.dataService, initial.props.dataService); assert.equal(creations.length, count);
  part.properties.backendBaseUrl = "https://inactive.example.com"; part.properties.backendResourceId = "inactive"; part.render();
  assert.equal(current().key, initial.key); assert.equal(current().props.dataService, initial.props.dataService);
  part.properties.siteUrl = "https://tenant.sharepoint.com/sites/other/"; part.render();
  assert.notEqual(current().key, initial.key); assert.notEqual(current().props.roleResolver, initial.props.roleResolver);
  assert.equal(current().props.connection.siteUrl, "https://tenant.sharepoint.com/sites/other");
  part.properties.dataSourceMode = "backend"; part.properties.backendBaseUrl = "https://api.example.com/api/v1/"; part.properties.backendResourceId = "api://app"; part.render();
  const backend = current();
  assert.equal(creations.filter(x => x.instance === backend.props.dataService)[0].kind, "backend");
  assert.equal(creations.filter(x => x.instance === backend.props.roleResolver)[0].kind, "backend-role");
  const client = creations.filter(x => x.kind === "client").at(-1)!;
  assert.equal(client.args[1], "api://app"); assert.equal(client.args[2], "https://api.example.com/api/v1");
  assert.equal(creations.filter(x => x.instance === backend.props.dataService)[0].args[0], client.instance);
  assert.equal(creations.filter(x => x.instance === backend.props.roleResolver)[0].args[0], client.instance);
  part.render(); assert.equal(current().props.dataService, backend.props.dataService);
  part.properties.siteUrl = "https://tenant.sharepoint.com/sites/inactive"; part.properties.betaAccessMode = true; part.render();
  assert.equal(current().key, backend.key); assert.equal(current().props.dataService, backend.props.dataService);
  part.properties.backendResourceId = "new-app"; part.render(); assert.notEqual(current().key, backend.key);
  const changed = current(); part.context.pageContext.user.email = "newcaller@example.com"; part.render(); assert.notEqual(current().key, changed.key);
});

test("local direct mode retains its explicit mock/admin resolver and explicit backend mode never falls back", () => {
  const part = webpart(true); part.render();
  assert.equal(creations.filter(x => x.instance === current().props.dataService)[0].kind, "mock");
  assert.equal(creations.filter(x => x.instance === current().props.roleResolver)[0].kind, "local-role");
  part.properties.dataSourceMode = "backend"; part.render();
  assert.equal(creations.filter(x => x.instance === current().props.dataService)[0].kind, "backend");
  assert.equal(creations.filter(x => x.instance === current().props.roleResolver)[0].kind, "backend-role");
});
