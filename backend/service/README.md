# Q-Star Issue Manager — backend service

Java/Spring Boot service scaffolded from time:matters' internal backend template
(`backend-template-develop`, Sept 2026 share from IT). Service-layer architecture: Controller →
Service → Repository, constructor-based dependency injection.

## Architecture: thin gateway in front of SharePoint

**SharePoint remains the actual data store** — the "Q-Star Issues" and "Q-Star Progress Log"
lists on the Quality site (see [`../sharepoint/qstar-sharepoint-graph-integration.md`](../sharepoint/qstar-sharepoint-graph-integration.md)).
This service sits in front of it as an authenticated gateway:

```
SPFx web part ──(Azure AD bearer token)──▶ this backend ──(Graph, app identity, Sites.Selected)──▶ SharePoint
```

- The frontend authenticates callers via Azure AD (`configuration/security/WebSecurityConfig.java`,
  OAuth2 resource server / JWT bearer).
- `infrastructure/sharepoint/SharePointGraphClient.java` reads/writes the SharePoint lists via
  Microsoft Graph, authenticating as this service's own app identity (client-credentials flow,
  `Sites.Selected` — not the signed-in user's own SharePoint permissions).
- The Postgres database wired up by the template (`QstarDatabase`, Flyway) exists only because
  the template's boot sequence expects one — there's no Q-Star business data in it. See
  `src/main/resources/db/migration/V1.0__Base.sql`.

## Naming (per the template's convention)

- `Order.java` → internal model. `OrderATO.java` → what's serialized to/from the frontend
  (generated from `api-contract/contract.yaml`). `OrderDAO.java` → JDBI data access.
- Here: `Issue.java` / `IssueATO` (generated) / no DAO — `IssueRepository.java` talks to Graph
  directly rather than JDBI, since SharePoint is the store, not Postgres.

## Building locally

Requires JDK 21 (this session used [Eclipse Temurin](https://adoptium.net/), installed to
`~/java` with no admin rights needed — a system package manager works too) and no other local
setup; Gradle wrapper handles the rest.

```bash
export JAVA_HOME=/path/to/jdk-21
./gradlew build          # runs OpenAPI codegen, compiles, packages the bootJar
./gradlew bootRun         # needs a real Postgres + the TM_QSTAR_* env vars below to actually start
```

`compileJava`/`build` succeed standalone (verified). `bootRun`/`test` need a live Postgres
instance and real Azure AD / SharePoint credentials that don't exist yet — see "What's still
needed" below.

### Required environment variables (`local-envs.ps1.template` → copy to `local-envs.ps1`)

| Variable | Purpose |
|---|---|
| `TM_QSTAR_DATASOURCE_URL/USERNAME/PASSWORD` | Postgres (template plumbing only, see above) |
| `TM_QSTAR_AD_CLIENT_ID/SECRET/TENANT_ID` | Validates bearer tokens from the SPFx web part |
| `TM_QSTAR_SP_CLIENT_ID/SECRET` | This service's own Graph app identity |
| `TM_QSTAR_SP_SITE_HOSTNAME/SITE_PATH` | The Quality SharePoint site, e.g. `contoso.sharepoint.com` / `/sites/Quality` |
| `TM_QSTAR_BUGSNAG_KEY` | Error tracking (optional) |

The `TM_QSTAR_AD_*` and `TM_QSTAR_SP_*` credentials may end up being the same Entra app
registration (granted both an exposed API scope and `Sites.Selected` application permissions) or
two separate ones — either way an admin needs to create it/them and grant `Sites.Selected` write
access to the Quality site (see the integration doc, §4.5).

## API

Defined in [`api-contract/contract.yaml`](api-contract/contract.yaml) (OpenAPI 3.0), generated
into Java at build time — `IssuesApi`/`SettingsApi`/`DiagnosticsApi` interfaces plus `*ATO`
models. Endpoints: `GET/PATCH/POST /issues*`, `GET/PUT /settings`, `GET /diagnostics` (this
service's own SharePoint connectivity self-test — see `DiagnosticsService.java`).

## What was fixed in the template

This template zip hadn't been compiled end-to-end since being upgraded to Spring Boot 3.3.4 —
building it surfaced several real gaps, all fixed here (not just worked around):

- **Missing dependencies**: `com.remondis:remap` (the mapper classes' actual library — only
  MapStruct was wired up, unused by any example code), `com.bugsnag:bugsnag`/`bugsnag-spring`,
  `spring-boot-starter-validation`.
- **`javax.*` → `jakarta.*`**: Spring Boot 3 requires Jakarta EE; the custom OpenAPI codegen
  templates (`openApiTemplate/*.mustache`) and several hand-written files (`ApiUtil.java`,
  `CustomExceptionHandler.java`, `CustomErrorController.java`) still imported `javax.validation`/
  `javax.servlet`. Also added the generator's own `useJakartaEe: true` option.
- **`AadResourceServerWebSecurityConfigurerAdapter` doesn't exist** in the
  `spring-cloud-azure-starter-active-directory:5.17.1` version build.gradle declares (confirmed
  by inspecting the actual jar) — rewritten against the real class,
  `AadResourceServerHttpSecurityConfigurer`, in the modern Spring Security 6 `SecurityFilterChain`
  bean style rather than the removed `WebSecurityConfigurerAdapter` inheritance style.
  `UserRepository.java` similarly referenced the legacy `com.microsoft.azure.spring...UserPrincipal`
  API instead of the `Jwt` principal the resource-server setup actually provides.
  `AbstractErrorController`/`ErrorController.getErrorPath()` no longer exist either in Spring Boot
  3 — `ErrorController` is now an empty marker interface.
  `ResponseEntityExceptionHandler`'s overridable methods changed their status parameter from
  `HttpStatus` to `HttpStatusCode`.
- **Missing `interfaceOnly: true`**: without it, the generator also emits a concrete
  `*ApiController` stubbed to `NOT_IMPLEMENTED` on every method, which collides with a
  hand-written controller of the same name/interface once generated sources actually feed into
  the build (they didn't before either — see next point). That's why the template's own
  `SampleApiController.java` had to be a manually reconciled copy rather than something
  regenerated on every build.
- **Generated sources never reached the compiled classpath**: `openApiGenerate` wrote to
  `build/generated` but nothing added that to `sourceSets` or ran it before `compileJava` — wired
  up properly so the `Api` interfaces/`*ATO` models regenerate from the contract on every build.
- **Hardcoded leftovers from a different real project** ("globaloffer", evidently what this
  template was extracted from): a dead import in `model.mustache`
  (`com.timematters.globaloffer.model.AbstractModel` — turned out `AbstractModel` itself is
  real/load-bearing, just needed providing, see `model/AbstractModel.java`), a malformed line and
  an unreferenced duplicate line in `local-envs.ps1.template`, a hardcoded property key
  (`tm.go.bugsnag.api.key`), and stale config-map/port-name entries in `kubernetes.yaml`
  referencing another project's files.
- **`Error`/`ValidationError` response shape mismatch**: the original contract split these into
  two schemas, but `CustomExceptionHandler`/`ValidationError.toErrorResponse()` always return
  `ErrorATO` with a free-form `meta` bag — the contract's `ValidationError` response now points at
  the same `Error` schema to match what the code actually returns.

Full details and rationale are in code comments at each fix site, and in
[`openApiTemplate/README.md`](openApiTemplate/README.md) for the codegen template fixes
specifically.

## What's still needed (not something I can do without your infra/access)

1. **A real Postgres instance**, even though nothing Q-Star-specific lives in it yet.
2. **Two Entra app registrations** (or one doing both jobs) — one for the SPFx web part to
   request tokens against this API, one for this API's own Graph client-credentials calls with
   `Sites.Selected` granted on the Quality site.
3. **An Azure DevOps project** to actually run `azure-pipelines.yml`, and access to
   `tmregistry.azurecr.io` (the `Dockerfile`'s base image is private to time:matters) — I can't
   build/push the Docker image or deploy to Kubernetes from here.
4. **A Kubernetes namespace** for `kubernetes.yaml`'s Deployment/Service/Ingress once the above
   exist.

Once 1–2 exist, `./gradlew test`/`bootRun` become runnable and worth trying before 3–4.
