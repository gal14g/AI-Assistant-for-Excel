# Deploying AI Assistant For Excel to OpenShift

End-to-end guide to deploying the add-in to an OpenShift cluster behind a
TLS-terminated Route. Takes ~15 minutes first time through.

---

## 0. Prerequisites

- Access to an OpenShift cluster + `oc` CLI logged in to the target project
- A container registry you can push to (GitLab CR, Quay, ECR, etc.)
- An LLM API key (OpenAI / Anthropic / Cohere / Gemini — or a reachable Ollama
  endpoint if running fully air-gapped)
- Users will install the Excel add-in via the manifest URL — they need network
  access to the Route hostname from their Excel client

---

## 1. Decide your Route hostname

The hostname is baked into `manifest.xml` at **build time** via the
`FRONTEND_URL` docker build-arg. It must be the **exact** public URL users will
hit from Excel. A mismatch here and Office rejects the add-in.

Pick one of:

- **A:** Let OpenShift auto-assign (e.g. `excel-assistant-myproject.apps.cluster.example.com`)
  — run the route-create step first, capture the hostname, then build.
- **B:** Reserve a host up front (e.g. `excel-assistant.apps.cluster.example.com`)
  — set it in `openshift/route.yaml` before applying.

Export it for the rest of this guide:

```bash
export FRONTEND_URL="https://excel-assistant.apps.your-cluster.example.com"
```

---

## 2. Build and push the image

```bash
# From the repo root
docker build \
  --build-arg FRONTEND_URL="$FRONTEND_URL" \
  -t your-registry.example.com/your-group/excel-assistant:v1.1.0 \
  .

docker push your-registry.example.com/your-group/excel-assistant:v1.1.0
```

The build is multi-stage:

1. Node builds the React add-in with `FRONTEND_URL` baked into `manifest.xml`
2. Python installs backend dependencies into a venv
3. Final slim image = backend + `./static/` (built frontend) + bundled
   `paraphrase-multilingual-MiniLM-L12-v2` embedding model (~420 MB, multilingual — Hebrew + 50 languages). No network calls at runtime.

If your cluster can't reach your registry directly, configure an image pull
secret:

```bash
oc create secret docker-registry regcred \
  --docker-server=your-registry.example.com \
  --docker-username=... \
  --docker-password=... \
  --docker-email=...
oc secrets link default regcred --for=pull
```

---

## 3. Edit the manifests for your cluster

Before applying, update these placeholders:

| File | Field | Change to |
|---|---|---|
| `openshift/deployment.yaml` | `image:` | Your pushed image tag |
| `openshift/route.yaml` | `host:` (uncomment) | Your hostname, matching `FRONTEND_URL` |
| `openshift/configmap.yaml` | `CORS_ORIGINS` | `'["https://your-host.example.com"]'` |

The `CORS_ORIGINS` value must be a JSON array string. Use the same hostname as
`FRONTEND_URL`. In OpenShift the frontend is same-origin so CORS is technically
not required for the Office.js iframe, but setting it correctly prevents
surprises if you add a separate frontend origin later.

---

## 4. Create the secret

Never commit the API key. Create it directly:

```bash
oc create secret generic excel-assistant-secrets \
  --from-literal=LLM_API_KEY='sk-...your-key...'
```

For Azure OpenAI or other providers needing extra vars, add them here:

```bash
oc create secret generic excel-assistant-secrets \
  --from-literal=LLM_API_KEY='...' \
  --from-literal=LLM_API_VERSION='2024-02-15-preview'
```

---

## 5. Apply manifests

```bash
# Storage first (PVC must exist before the Deployment references it)
oc apply -f openshift/pvc.yaml

# Config (already-created secret stays untouched)
oc apply -f openshift/configmap.yaml

# App + networking
oc apply -f openshift/deployment.yaml
oc apply -f openshift/service.yaml
oc apply -f openshift/route.yaml
```

Or in one shot (after the secret exists):

```bash
oc apply -f openshift/
```

---

## 6. Verify

```bash
# Pod should reach Running, 1/1 ready
oc get pods -l app=excel-assistant -w

# Readiness endpoint — should return {"ready": true, "model": "..."}
oc exec deploy/excel-assistant -- curl -s http://localhost:8080/ready

# Route URL — should match your FRONTEND_URL exactly
oc get route excel-assistant -o jsonpath='{.spec.host}{"\n"}'

# Fetch manifest.xml through the Route
curl -sfI "$FRONTEND_URL/manifest.xml"
```

The first pod start takes 20-40 s because ChromaDB indexes the 51 capabilities
into the PVC. Subsequent restarts are fast (index is persisted).

If `/ready` returns 503 with `"LLM_API_KEY is not set"`, the secret wasn't
mounted — check that the secret name matches `excel-assistant-secrets` exactly.

---

## 7. Install the add-in in Excel

Users sideload from the manifest URL:

- **Excel Desktop:** Insert → My Add-ins → Upload My Add-in → point at
  `https://your-host.example.com/manifest.xml`
- **Excel Web:** Insert → Office Add-ins → Upload My Add-in → same URL
- **Org-wide deployment:** Microsoft 365 Admin Center → Integrated apps →
  Upload custom apps → host the manifest at the same URL

---

## 8. Updating

```bash
# Rebuild and push with a new tag
docker build --build-arg FRONTEND_URL="$FRONTEND_URL" \
  -t your-registry.example.com/your-group/excel-assistant:v1.2.0 .
docker push your-registry.example.com/your-group/excel-assistant:v1.2.0

# Roll the deployment
oc set image deploy/excel-assistant excel-assistant=your-registry.example.com/your-group/excel-assistant:v1.2.0
oc rollout status deploy/excel-assistant
```

RollingUpdate strategy keeps the old pod up until the new one passes its
readiness probe (zero downtime).

---

## 9. Scaling constraints

**Do not scale replicas above 1.** The app uses:

- **SQLite** for feedback/conversations (single file on the PVC)
- **ChromaDB PersistentClient** for vector collections (single file on the PVC)
- **ReadWriteOnce** PVC — only one pod can mount it

To run multiple replicas you must migrate to Postgres + pgvector — see the
"Migrating to a production database" section in `README.md`. The table schemas
stay identical; only the storage backends change.

Vertical scaling (more CPU/memory per pod) works fine within the deployment's
existing limits (2 CPU / 2 Gi). Uvicorn is configured with `--workers 1` to
avoid loading the embedding model twice into memory.

---

## 10. Troubleshooting

| Symptom | Likely cause |
|---|---|
| Pod CrashLoopBackoff, logs show "LLM_API_KEY" | Secret not created or wrong name |
| Office.js: "Can't load add-in" | `FRONTEND_URL` build-arg ≠ Route hostname |
| `/ready` returns 503 | Check the `errors` array in the JSON response |
| Startup probe failing after 24 tries | PVC permissions — check `oc describe pod` for mount errors |
| Slow first response only | ChromaDB loading — warm up by hitting `/ready` after rollout |
| 502 from Route on large plans | Increase `haproxy.router.openshift.io/timeout` in `route.yaml` |

Useful commands:

```bash
# Full startup logs
oc logs deploy/excel-assistant --tail=200

# Describe pod (events, mount errors, probe failures)
oc describe pod -l app=excel-assistant

# Shell into the running pod
oc exec -it deploy/excel-assistant -- /bin/bash

# Check PVC contents (data/chroma/ and data/feedback.db should exist)
oc exec deploy/excel-assistant -- ls -la /app/data
```
