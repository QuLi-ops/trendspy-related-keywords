#!/usr/bin/env bash
set -euo pipefail

REPO_DIR="/home/ubuntu/trendspy-related-keywords"
BRANCH="main"
REMOTE="origin"
SERVICE_NAME="trends-monitor.service"
LOCK_FILE="/tmp/trendspy-related-keywords-auto-update.lock"

log() {
  printf '[%s] %s\n' "$(date -Is)" "$*"
}

run_git_as_owner() {
  if [ "$(id -u)" -eq 0 ]; then
    sudo -u ubuntu git -C "$REPO_DIR" "$@"
  else
    git -C "$REPO_DIR" "$@"
  fi
}

(
  flock -n 9 || {
    log "Another update check is already running; skipping."
    exit 0
  }

  cd "$REPO_DIR"

  current_branch="$(run_git_as_owner rev-parse --abbrev-ref HEAD)"
  if [ "$current_branch" != "$BRANCH" ]; then
    log "Current branch is $current_branch, expected $BRANCH; skipping."
    exit 0
  fi

  if [ -n "$(run_git_as_owner status --porcelain)" ]; then
    log "Working tree has local changes; skipping auto update."
    run_git_as_owner status --short
    exit 0
  fi

  log "Fetching $REMOTE/$BRANCH..."
  run_git_as_owner fetch "$REMOTE" "$BRANCH"

  local_sha="$(run_git_as_owner rev-parse "$BRANCH")"
  remote_sha="$(run_git_as_owner rev-parse "$REMOTE/$BRANCH")"

  if [ "$local_sha" = "$remote_sha" ]; then
    log "Already up to date at $local_sha."
    exit 0
  fi

  if ! run_git_as_owner merge-base --is-ancestor "$local_sha" "$remote_sha"; then
    log "Remote is not a fast-forward from local; skipping."
    exit 1
  fi

  old_requirements_sha=""
  if [ -f requirements.txt ]; then
    old_requirements_sha="$(sha256sum requirements.txt | awk '{print $1}')"
  fi

  log "Fast-forwarding $local_sha -> $remote_sha..."
  run_git_as_owner pull --ff-only "$REMOTE" "$BRANCH"

  new_requirements_sha=""
  if [ -f requirements.txt ]; then
    new_requirements_sha="$(sha256sum requirements.txt | awk '{print $1}')"
  fi

  if [ "$old_requirements_sha" != "$new_requirements_sha" ] && [ -x venv/bin/pip ]; then
    log "requirements.txt changed; installing dependencies..."
    sudo -u ubuntu venv/bin/pip install -r requirements.txt
  fi

  log "Restarting $SERVICE_NAME..."
  systemctl restart "$SERVICE_NAME"
  systemctl --no-pager --quiet is-active "$SERVICE_NAME"
  log "Update complete; $SERVICE_NAME is active."
) 9>"$LOCK_FILE"
