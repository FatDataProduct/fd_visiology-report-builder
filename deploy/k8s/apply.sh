#!/usr/bin/env bash
# Apply non-secret GitOps manifests. Usage: ./deploy/k8s/apply.sh main
# Does not apply Secret.
set -euo pipefail
ENV_TARGET="${1:?usage: $0 main}"
ROOT="$(cd "$(dirname "$0")" && pwd)"
NS="front"
CTX="front"
case "$ENV_TARGET" in
  main) ;;
  *) echo "usage: $0 main" >&2; exit 1 ;;
esac
echo "Applying ${ENV_TARGET} manifests to namespace/${NS} context/${CTX}"
kubectl --context="${CTX}" apply -n "${NS}" -f "${ROOT}/${ENV_TARGET}/deployment.yaml"
kubectl --context="${CTX}" apply -n "${NS}" -f "${ROOT}/${ENV_TARGET}/service.yaml"
kubectl --context="${CTX}" apply -n "${NS}" -f "${ROOT}/${ENV_TARGET}/httproute.yaml"
echo "Applied non-secret manifests for ${ENV_TARGET} in ns/${NS}. Secrets are out of band."
