#!/usr/bin/env bash
# Compute the next semantic version from Conventional Commits made since the last
# release tag. The latest vMAJOR.MINOR.PATCH tag is the source of truth for the
# current released version; this script decides how far to bump it.
#
# Output (stdout):
#   "<major>.<minor>.<patch> <level>"  when a release is warranted
#   "none"                             when no commit since the last release
#                                      warrants one (only chore/docs/ci/etc.)
#
# Bump rules (https://www.conventionalcommits.org + SemVer):
#   type!:  or  BREAKING CHANGE: footer  -> major
#   feat:                                -> minor
#   fix:                                 -> patch
#   everything else                      -> no bump
#
# Precedence is highest-wins: one breaking commit forces a major even if there
# are also feat/fix commits in the same range.
set -euo pipefail

# Latest release tag shaped vMAJOR.MINOR.PATCH. The git glob is a coarse first
# pass (it also matches vendor-libs-* and prerelease tags like v2.0.0-rc1); the
# grep anchors an exact vX.Y.Z so a prerelease tag never becomes the base and
# corrupts the arithmetic below. --sort=-v:refname orders newest semver first.
last_tag=$(git tag --list 'v[0-9]*.[0-9]*.[0-9]*' --sort=-v:refname \
  | grep -E '^v[0-9]+\.[0-9]+\.[0-9]+$' \
  | head -n1 || true)

if [ -z "$last_tag" ]; then
  base="0.0.0"
  range=""            # no prior release: consider the whole history
else
  base="${last_tag#v}"
  range="${last_tag}..HEAD"
fi

IFS='.' read -r major minor patch <<EOF
$base
EOF

level="none"

# Walk every commit in the range. %B is the full message (subject + body); the
# NUL terminator keeps multi-line bodies grouped so BREAKING CHANGE footers survive.
while IFS= read -r -d '' msg; do
  # git log puts a newline between records, so every record after the first
  # starts with one; without stripping it the subject would read as empty.
  msg="${msg#$'\n'}"
  [ -n "$msg" ] || continue
  subject="${msg%%$'\n'*}"

  if printf '%s' "$subject" | grep -Eq '^[a-z]+(\([^)]+\))?!:' \
     || printf '%s\n' "$msg" | grep -Eq '^BREAKING[ -]CHANGE:'; then
    level="major"
    break
  fi

  if printf '%s' "$subject" | grep -Eq '^feat(\([^)]+\))?:'; then
    level="minor"
  elif printf '%s' "$subject" | grep -Eq '^fix(\([^)]+\))?:'; then
    [ "$level" = "minor" ] || level="patch"
  fi
done < <(git log ${range:+$range} --format='%B%x00')

case "$level" in
  major) major=$((major + 1)); minor=0; patch=0 ;;
  minor) minor=$((minor + 1)); patch=0 ;;
  patch) patch=$((patch + 1)) ;;
  none)  echo "none"; exit 0 ;;
esac

echo "${major}.${minor}.${patch} ${level}"
