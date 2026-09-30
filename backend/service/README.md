# Q-Star Java backend

This optional Java 21 / Spring Boot 3.3.4 API uses SharePoint as the data store.
It is **disabled by default**. Direct SharePoint remains the SPFx default mode.
Disabled startup and local tests require no tenant credentials, database, consent changes,
container registry access or external telemetry.

## Permission and activation boundary

Do not enable this backend until IT verifies a **narrowly scoped delegated SharePoint REST
permission grant** and completes the tenant acceptance checks below. The repository's
single-site permission requirement still applies. A configured site URL limits this code's
requests; it does not narrow an Entra permission grant.

Microsoft documents delegated `Sites.Selected` for Graph, but that is not evidence that the
same delegated permission is available for every SharePoint REST operation used here.
The REST/CSOM guidance has different permission requirements. Do not promise or assume that
Graph consent authorizes REST. If a narrow delegated REST configuration cannot be established,
keep backend mode disabled and continue using the repaired direct SharePoint service.
`AllSites.Write` is a broader delegated alternative, **not an approved setup step**: it requires
an explicit user/IT exception to the single-site requirement. There is no automatic permission
broadening, application-token fallback, or production consent script in this service.

Sources: [selected permissions](https://learn.microsoft.com/en-us/graph/permissions-selected-overview),
[delegated Sites.Selected announcement](https://devblogs.microsoft.com/microsoft365dev/sharepoint-now-supports-delegated-sites-selected-authentication/),
[SharePoint REST/CSOM distinction](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins-modernize/use-remote-event-receivers-without-azure-acs-dependency).

## Runtime identity and authorization

```text
SPFx -- API-audience delegated bearer --> Java API
Java API -- OBO, signed-in caller's SharePoint token --> pinned SharePoint site
```

The resource server validates the JWT signature, timestamps, exact tenant/issuer and API
audience. It requires the configured delegated scope (default `user_impersonation`) and a
user object ID. App-only tokens are rejected. App-role values are an explicit allowlist;
unknown or missing roles are read-only. Email and display name are never authorization keys.
Raw group claims are not used, so group overage cannot grant a fallback role.

`CurrentUserProvider` combines the verified API role with the delegated
`/_api/web/currentuser` numeric identity, name and email. `/api/v1/me` returns this identity
and the actual configured connection. The frontend must use this response in backend mode,
not infer backend privileges from direct SharePoint groups or a mock user.

| API capability | Required role and additional checks |
|---|---|
| Read issues/settings, `/me` | Valid delegated API token and native SharePoint access |
| Create issues | Admin or QM, plus native SharePoint write access |
| Update issues / append progress | Admin/QM, or current owner with permitted owner fields/transitions |
| Save settings / run diagnostics | Admin; SharePoint still enforces the caller's list permissions |

Role names do not override native SharePoint permissions. Owner checks use the freshly read
issue's `TaskOwnerId` and the authoritative numeric caller ID. Journals append only in the
secured `issue-<ID>` folder, and SharePoint supplies native `Author`/`Created`; caller-provided
author, timestamps and parent metadata do not establish identity or journal membership.
Keep provisioning and Flow C's item/folder permission reconciliation in place.

The OBO client requests only `https://<configured-sharepoint-host>/.default`, using this API's
confidential-client credential and the incoming caller assertion. The JWT-bearer provider is
called directly and caches the resulting token only within the HTTP request. It does not use
a global application token, an authorization-code session, or Graph as an alternate data path.
The API never returns its downstream token. Consent/Conditional Access failures fail closed;
affected callers may need to sign in again. Tenant acceptance must cover the organization's
Conditional Access policy before enabling this mode.

Microsoft references: [OBO flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-on-behalf-of-flow),
[claim validation](https://learn.microsoft.com/en-us/entra/identity-platform/claims-validation),
[Spring Cloud Azure security](https://learn.microsoft.com/en-us/azure/developer/java/spring-framework/spring-security-support).
The implementation targets Azure Spring **5.17.1** and its
[`spring.cloud.azure.active-directory` configuration](https://github.com/Azure/azure-sdk-for-java/blob/spring-cloud-azure_5.17.1/sdk/spring/spring-cloud-azure-autoconfigure/src/main/java/com/azure/spring/cloud/autoconfigure/implementation/aad/configuration/properties/AadAuthenticationProperties.java).
It adds stricter tenant/issuer validation than that release's default validator.

## Configuration

Copy `local-envs.ps1.template` to a private ignored file if PowerShell is used. Keep secrets
in the deployment secret store; do not put them in webpart properties, package files or Git.
The API can use an OBO client secret or a registered PFX/P12 certificate. If a certificate path
is configured, certificate authentication takes precedence. The mounted private key and
password must remain accessible only to the service identity.

| Environment variable | Meaning |
|---|---|
| `TM_QSTAR_BACKEND_ENABLED` | Default `false`; functional endpoints return 503 while disabled |
| `TM_QSTAR_AD_TENANT_ID` | One tenant GUID; `common`/`organizations` are rejected |
| `TM_QSTAR_AD_CLIENT_ID` | This API's application GUID and v2 audience |
| `TM_QSTAR_AD_APP_ID_URI` | Optional exact v1 API audience, e.g. `api://<GUID>` |
| `TM_QSTAR_API_SCOPE` | Required delegated API scope; default `user_impersonation` |
| `TM_QSTAR_AD_CLIENT_SECRET` | OBO confidential-client secret, if certificate auth is not used |
| `TM_QSTAR_AD_CERTIFICATE_PATH/PASSWORD` | OBO certificate file and private-key password |
| `TM_QSTAR_ALLOWED_ORIGINS` | Comma-separated exact origins, normally `https://<tenant>.sharepoint.com` |
| `TM_QSTAR_SP_SITE_URL` | One absolute HTTPS SharePoint site URL |
| `TM_QSTAR_SP_SITE_HOSTNAME/SITE_PATH` | Legacy input, used only if `SITE_URL` is empty |
| `TM_QSTAR_ISSUES_LIST/PROGRESS_LIST/CONFIG_LIST` | Defaults: Q-Star Issues / Q-Star Progress Log / Q-Star Config |
| `TM_QSTAR_BETA_ACCESS_MODE` | Connection/profile description only; never relaxes API authorization |
| `TM_QSTAR_ADMIN_ROLE/QM_ROLE/OWNER_ROLE/READER_ROLE` | Distinct app-role values; defaults `QStar.Admin`, `QStar.QM`, `QStar.Owner`, `QStar.Reader` |
| `TM_QSTAR_BUGSNAG_ENABLED` | Default `false`; separate explicit telemetry opt-in |
| `TM_QSTAR_BUGSNAG_KEY` | Required only when Bugsnag is deliberately enabled |

CORS accepts exact origins, not wildcard tenant or hostname suffix patterns. Explicit
localhost HTTP origins are available for local browser development. Production configuration
should contain only approved HTTPS origins. Bearer requests do not require cookies.
`ETag`, `Location`, `X-QStar-Reference` and `X-QStar-Entry-Id` are exposed so clients can
recover an accepted operation without blindly repeating a POST.

The old template's `MicrosoftGraphClient`, `UserRepository`, `UserService`, `QstarDatabase`
and database migration examples remain unwired scaffolding. None is an active Spring service.
Database/Flyway autoconfiguration is excluded because Q-Star does not use it. Reintroducing a
database or app-only access requires a separate design change. HTTP body logging, Bugsnag and
the template New Relic agent are not activated by default.

## Local build and validation

From `backend/service`, with JDK 21 on `JAVA_HOME`:

```bash
bash ./gradlew --no-daemon test
bash ./gradlew --no-daemon build
bash ./gradlew --no-daemon bootRun
```

With the default configuration, `bootRun` starts a disabled API; `/api/v1/me` returns
`503 BACKEND_DISABLED`. Tests use local mocks and synthetic credentials. They do not contact
an Entra tenant or SharePoint, grant consent, run a provisioning script, publish an image,
or deploy a flow/container. The suite covers role and ownership boundaries, delegated token
requests, native journal mapping, ETag conflict handling, accepted-write receipts, CORS and
default-disabled startup. These checks do not establish live tenant compatibility.

The REST contract is generated from `api-contract/contract.yaml` during compilation.
Main endpoints use `/api/v1`: `GET /me`, `GET/POST /issues`, `GET/PATCH /issues/{id}`,
`POST /issues/{id}/progress`, `GET/PUT /settings`, and `GET /diagnostics`.
PATCH requires the caller's original `If-Match` ETag; 412 means reload before retrying.
Accepted writes can return a saved receipt with a readback warning. Clients must keep the
accepted identity and reconcile it rather than resubmit the mutation.

## IT acceptance before activation

1. Preserve the single-site requirement. Verify the actual delegated **SharePoint REST**
   scopes exposed by the tenant and their behavior for current-user lookup, list operations,
   Person resolution, and folder-based append. Leave the backend disabled if a narrow grant
   cannot satisfy these operations; seek an explicit decision before any broader proposal.
2. Expose the API delegated scope, authorize the SPFx client through the tenant's normal
   approval process, and configure API-audience token acquisition. Define the four app roles
   for Users/Groups and assign them deliberately. Do not map group names into the roles claim.
3. Configure the API credential, exact tenant/audience, approved origin and pinned site/list
   names. Keep the existing production item/folder ACLs and immutable reference-offset setup.
4. With separate Reader, Owner A, Owner B, QM and Admin accounts, verify denied cross-owner
   writes, restricted owner fields, settings administration, reassignment/removal revocation,
   and journal `Author`/`Created` fidelity. Confirm invalid issuer/audience/scope and application
   tokens are rejected, and Conditional Access failures do not trigger application fallback.
5. Verify a SharePoint-hosted SPFx page can make the authenticated CORS request. Exercise
   stale ETags and an accepted write followed by failed readback without duplicate creation.
6. Only after that acceptance, enable `TM_QSTAR_BACKEND_ENABLED` and select backend mode on
   the intended webpart. Keep direct SharePoint available for the agreed coexistence period.

No tenant acceptance, consent or deployment is performed by a local build or this repair.

## Deployment artifacts

`azure-pipelines.yml` uses JDK 21 and builds in `backend/service`, with the matching Docker
build context. It requires IT's Docker registry service connection and a valid `host`
pipeline variable. The pipeline can build/push its image and publish a rendered manifest;
it does not apply Kubernetes resources. Do not run it as part of local-only validation.

`kubernetes.yaml` uses `networking.k8s.io/v1`, an explicit Prefix `/api` route and the declared
service port. It preserves `/api/v1/...` when forwarding to Spring. TLS, ingress-nginx,
namespace, secret/config-map contents and the private base image must be supplied by IT.
The JVM uses a percentage of the container memory rather than a heap larger than its limit.
The private base image, registry access and live cluster behavior remain deployment checks.
