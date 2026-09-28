-- No Q-Star business tables here on purpose: SharePoint is the system of record
-- (see backend/sharepoint/qstar-sharepoint-graph-integration.md and
-- backend/service/README.md's "thin gateway" note). This Postgres database exists
-- only because the template's Flyway/JDBI wiring (QstarDatabase, spring.flyway.*)
-- expects one to be reachable at boot — a real instance still needs to be
-- provisioned even though nothing here stores Q-Star issue data. Add real tables
-- here if/when a genuine backend-local need comes up (e.g. an audit log of writes
-- proxied to SharePoint).
SELECT 1;
