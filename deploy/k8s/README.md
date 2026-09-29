# Kubernetes manifests (GitOps)

Namespace `front` на кластере `front`. Namespace уже есть, скрипт его не создаёт.
CI вызывает `kubectl --context=front`.

Живой контур в этом каталоге: `main`. `kubectl diff` по каждому файлу из `apply.sh` был пустой.

HTTPRoute hostname: reportbuilder.fatdataseo.com.

Gateway не создаётся здесь. Берётся уже существующий listener.

## Секреты

`apply.sh` секреты не применяет.
Своего Secret у Deployment нет.

## Что не входит в apply

Test-деплоя нет. Живой контур — ветка main.

## Apply

Скрипт рассчитан на SSH-хост CI, где есть контекст `front`:

```bash
bash deploy/k8s/apply.sh main
```
