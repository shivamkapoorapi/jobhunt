#!/usr/bin/env bash
# Push every variable from .env.vercel into Vercel (production + preview).
# Fallback for when the browser dialog will not save.
#   bash vercel_env.sh
set -u
command -v vercel >/dev/null 2>&1 || { echo "Installing Vercel CLI..."; npm i -g vercel || exit 1; }
vercel whoami >/dev/null 2>&1 || vercel login || exit 1
[ -d .vercel ] || vercel link || exit 1

while IFS= read -r line; do
  case "$line" in ''|\#*) continue ;; esac
  key=${line%%=*}
  val=${line#*=}
  [ -z "$key" ] || [ -z "$val" ] && continue
  for env in production preview; do
    vercel env rm "$key" "$env" --yes >/dev/null 2>&1   # replace if it exists
    printf '%s' "$val" | vercel env add "$key" "$env" >/dev/null 2>&1 \
      && echo "  set  $key ($env)" || echo "  FAIL $key ($env)"
  done
done < .env.vercel

echo
echo "Redeploying so the new variables take effect..."
vercel --prod
