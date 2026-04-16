#!/bin/bash
# ═══════════════════════════════════════════════════════════
# ProjectFlow™ — Teams App Packager
# Creates a Teams-ready .zip with your HTTPS tunnel URL
# ═══════════════════════════════════════════════════════════

set -e

TEAMS_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$TEAMS_DIR")"

echo "╔═══════════════════════════════════════════════╗"
echo "║   ProjectFlow™ — Teams App Packager           ║"
echo "╚═══════════════════════════════════════════════╝"
echo ""

# Check if URL argument provided
if [ -z "$1" ]; then
    echo "Usage: ./package.sh <YOUR_HTTPS_URL>"
    echo ""
    echo "Example:"
    echo "  ./package.sh https://my-projectflow.loca.lt"
    echo "  ./package.sh https://abc123.ngrok-free.app"
    echo "  ./package.sh https://projectflow.azurewebsites.net"
    echo ""
    echo "To get a free HTTPS tunnel, run in another terminal:"
    echo "  npx -y localtunnel --port 5173"
    echo ""
    exit 1
fi

DOMAIN="$1"
# Strip protocol for validDomains
DOMAIN_CLEAN=$(echo "$DOMAIN" | sed 's|https://||' | sed 's|http://||' | sed 's|/$||')

echo "→ Using domain: $DOMAIN_CLEAN"

# Create manifest with actual domain
cat > "$TEAMS_DIR/manifest.json" <<EOF
{
  "\$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.16/MicrosoftTeams.schema.json",
  "manifestVersion": "1.16",
  "version": "1.0.0",
  "id": "cb85d5a6-ed15-4a0e-8ead-40c6d45491a6",
  "developer": {
    "name": "Ahmed M. Fawzy",
    "websiteUrl": "https://${DOMAIN_CLEAN}",
    "privacyUrl": "https://${DOMAIN_CLEAN}",
    "termsOfUseUrl": "https://${DOMAIN_CLEAN}"
  },
  "name": {
    "short": "ProjectFlow",
    "full": "ProjectFlow™ — Professional Project Management"
  },
  "description": {
    "short": "Professional project management with Gantt, CPM, EVM & Portfolio reporting",
    "full": "ProjectFlow™ is a comprehensive, browser-based project management system featuring Gantt charts, Critical Path Method (CPM), Earned Value Management (EVM), resource leveling, portfolio dashboards, and Microsoft Planner integration. Built for teams that need enterprise-grade project controls."
  },
  "icons": {
    "color": "color.png",
    "outline": "outline.png"
  },
  "accentColor": "#6366f1",
  "staticTabs": [
    {
      "entityId": "projectflow-home",
      "name": "ProjectFlow",
      "contentUrl": "https://${DOMAIN_CLEAN}/index.html",
      "websiteUrl": "https://${DOMAIN_CLEAN}/index.html",
      "scopes": ["personal"]
    }
  ],
  "permissions": ["identity"],
  "validDomains": [
    "${DOMAIN_CLEAN}"
  ],
  "showLoadingIndicator": false,
  "isFullScreen": false
}
EOF

echo "→ Manifest updated ✓"

# Package as zip
cd "$TEAMS_DIR"
rm -f ProjectFlow.zip
zip -j ProjectFlow.zip manifest.json color.png outline.png

echo ""
echo "╔═══════════════════════════════════════════════╗"
echo "║  ✅ Package ready: teams-app/ProjectFlow.zip  ║"
echo "╠═══════════════════════════════════════════════╣"
echo "║  Upload to Teams:                             ║"
echo "║  Apps → Upload an app → Upload a custom app   ║"
echo "╚═══════════════════════════════════════════════╝"
