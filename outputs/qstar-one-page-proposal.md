# Q-Star: stabilize now, pilot with one client, then decide on SaaS

## The opportunity

Q-Star already has a validated Quality-management workflow and a strong working prototype. Before using it operationally or offering it to clients, we should complete a focused technical cleanup and confirm the solution in a real Microsoft 365 tenant.

The goal is not to over-engineer the internal tool. It is to make it safe and dependable for the Quality team while creating a clean foundation for future consultancy deployments.

## Immediate priority: a safe internal release

Four areas need attention:

1. **Fix the data contract** — correct the SharePoint progress-author and Person-field implementations, add paging, and decide how evidence attachments are stored.
2. **Strengthen access control** — resolve roles through Microsoft Entra groups, apply least-privilege SharePoint permissions, and protect configuration settings.
3. **Protect record integrity** — enforce permitted lifecycle transitions, required fields, concurrency checks, safe QS numbering, and a dependable audit history.
4. **Prove operations** — provision every required list, run end-to-end tests in a real tenant, validate reminders, and document deployment and support procedures.

**Internal-release target:** the application works with real SharePoint data, roles behave as agreed, identifiers remain unique, and important history cannot be silently lost or overwritten.

## Three routes to a consultancy product

| Route | Strength | Limitation | Best use |
|---|---|---|---|
| **1. SharePoint-only package** | Fastest and lowest operational overhead | Business rules and fine-grained security remain difficult to enforce | Small, trusted Quality teams wanting a controlled implementation service |
| **2. Customer-hosted Q-Star API** | Strong customer isolation, enforceable roles and workflow rules, data remains in the customer environment | Requires a repeatable Azure deployment and upgrade process | First consultancy product and design-partner deployments |
| **3. Company-hosted SaaS** | Central onboarding, upgrades, telemetry, and subscription potential | We assume responsibility for tenant isolation, availability, compliance, incident response, backups, and customer data processing | Later-stage standardized product with proven repeat demand |

## Recommendation

**Use option 2 as the initial consultancy model.**

After stabilizing the internal release, package a small Entra-protected API that can be deployed inside each customer’s Azure/Microsoft environment. The customer retains its identities, SharePoint data, logs, and operational boundary. We retain the reusable product assets: source code, automated tests, deployment templates, upgrade packages, implementation guidance, and support offering.

This provides stronger governance than a SharePoint-only solution without immediately taking on the operational and compliance responsibilities of running a multi-customer SaaS platform.

## Proposed decision path

**Now — stabilize:** approve the internal cleanup scope and assign technical ownership.

**Next — pilot:** select one prospective customer as a design partner and deploy the customer-hosted model.

**Later — decide:** use evidence from the pilot—customer demand, common requirements, deployment effort, support burden, and commercial potential—to choose between repeat customer-hosted deployments and a centrally hosted SaaS product.

## Decision requested from Pierre

1. Approve the focused cleanup needed for a safe internal release.
2. Identify one suitable prospective client or design partner.
3. Postpone the SaaS commitment until the first consultancy deployment provides real evidence.

---

*Prepared from the Q-Star repository and architecture review, July 2026. Effort, ownership, and customer security requirements should be validated before implementation commitments are made.*
