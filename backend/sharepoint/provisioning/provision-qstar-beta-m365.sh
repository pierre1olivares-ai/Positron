#!/usr/bin/env bash
# =====================================================================
#  Q-Star Issue Manager — beta provisioning (CLI for Microsoft 365)
# =====================================================================
#  Creates only the Q-Star lists, fields, indexes, and native Person
#  columns. It deliberately does not create Q-Star groups or assign site
#  permissions. The web part must have Beta access mode enabled.
#
#  PREREQUISITES
#    npm i -g @pnp/cli-microsoft365
#    m365 login
#
#  RUN
#    SITE="https://contoso.sharepoint.com/sites/QStarBeta" ./provision-qstar-beta-m365.sh
# =====================================================================
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
export CREATE_ROLE_GROUPS=0
exec "$SCRIPT_DIR/provision-qstar-m365.sh"
