#!/usr/bin/env bash
# ═══════════════════════════════════════════════════════════════════════════════
# AI Assistant For Excel — Quick Deploy to OpenShift
#
# Usage:
#   ./openshift/deploy.sh <LLM_API_KEY> [IMAGE] [NAMESPACE]
#
# Example:
#   ./openshift/deploy.sh sk-ant-abc123
#   ./openshift/deploy.sh sk-ant-abc123 quay.io/myorg/excel-assistant:v1.1.0
#   ./openshift/deploy.sh sk-ant-abc123 quay.io/myorg/excel-assistant:v1.1.0 my-namespace
#
# Prerequisites:
#   - oc CLI installed and logged in (oc login ...)
#   - Docker image already built and pushed to a registry
# ═══════════════════════════════════════════════════════════════════════════════
set -euo pipefail

LLM_API_KEY="${1:?Usage: $0 <LLM_API_KEY> [IMAGE] [NAMESPACE]}"
IMAGE="${2:-registry.gitlab.com/YOUR_GROUP/excel-assistant:latest}"
NAMESPACE="${3:-}"

# Switch to namespace if provided
if [ -n "$NAMESPACE" ]; then
  oc project "$NAMESPACE"
fi

echo "==> Deploying AI Assistant For Excel..."

# 1. Create secret (skip if exists)
if oc get secret excel-assistant-secrets &>/dev/null; then
  echo "  Secret exists — updating LLM_API_KEY..."
  oc delete secret excel-assistant-secrets
fi
oc create secret generic excel-assistant-secrets \
  --from-literal=LLM_API_KEY="$LLM_API_KEY"

# 2. Apply all manifests
echo "  Applying manifests..."
oc apply -f openshift/pvc.yaml
oc apply -f openshift/configmap.yaml
oc apply -f openshift/deployment.yaml
oc apply -f openshift/service.yaml
oc apply -f openshift/route.yaml

# 3. Set the image
echo "  Setting image to: $IMAGE"
oc set image deployment/excel-assistant excel-assistant="$IMAGE"

# 4. Wait for rollout
echo "  Waiting for rollout..."
oc rollout status deployment/excel-assistant --timeout=180s

# 5. Get the route URL
ROUTE_URL=$(oc get route excel-assistant -o jsonpath='{.spec.host}' 2>/dev/null || echo "")
echo ""
echo "==> Deployment complete!"
if [ -n "$ROUTE_URL" ]; then
  echo ""
  echo "  App URL:      https://$ROUTE_URL"
  echo "  Health check: https://$ROUTE_URL/health"
  echo "  Manifest URL: https://$ROUTE_URL/manifest.xml"
  echo ""
  echo "  To install the add-in in Excel:"
  echo "    1. Open Excel"
  echo "    2. Insert > Add-ins > Upload My Add-in"
  echo "    3. Enter: https://$ROUTE_URL/manifest.xml"
  echo ""
  echo "  For org-wide deployment:"
  echo "    1. Go to admin.microsoft.com"
  echo "    2. Settings > Integrated apps > Upload custom apps"
  echo "    3. Enter manifest URL: https://$ROUTE_URL/manifest.xml"
  echo "    4. Assign to users/groups — rolls out within 24 hours"
fi
