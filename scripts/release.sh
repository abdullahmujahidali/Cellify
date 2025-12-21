#!/bin/bash

# Cellify Release Script
# Usage: ./scripts/release.sh [patch|minor|major]

set -e

RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

VERSION_TYPE=${1:-patch}

if [[ ! "$VERSION_TYPE" =~ ^(patch|minor|major)$ ]]; then
  echo -e "${RED}Error: Invalid version type. Use 'patch', 'minor', or 'major'${NC}"
  exit 1
fi

echo -e "${YELLOW}🚀 Starting Cellify release process...${NC}"

CURRENT_BRANCH=$(git branch --show-current)
if [ "$CURRENT_BRANCH" != "main" ]; then
  echo -e "${RED}Error: Must be on 'main' branch to release. Currently on '$CURRENT_BRANCH'${NC}"
  exit 1
fi

if [ -n "$(git status --porcelain)" ]; then
  echo -e "${RED}Error: Working directory is not clean. Commit or stash changes first.${NC}"
  exit 1
fi

echo -e "${GREEN}📥 Pulling latest changes...${NC}"
git pull origin main

echo -e "${GREEN}🧪 Running tests...${NC}"
npm test

echo -e "${GREEN}🔨 Building...${NC}"
npm run build

CURRENT_VERSION=$(node -p "require('./package.json').version")
echo -e "${GREEN}📌 Current version: ${CURRENT_VERSION}${NC}"

echo -e "${GREEN}📦 Bumping ${VERSION_TYPE} version...${NC}"
NEW_VERSION=$(npm version $VERSION_TYPE --no-git-tag-version)
NEW_VERSION=${NEW_VERSION#v} # Remove 'v' prefix if present

echo -e "${GREEN}📌 New version: ${NEW_VERSION}${NC}"

TODAY=$(date +%Y-%m-%d)
sed -i.bak "s/## \[Unreleased\]/## [Unreleased]\n\n## [${NEW_VERSION}] - ${TODAY}/" CHANGELOG.md

sed -i.bak "s|\[Unreleased\]: \(.*\)/compare/v.*\.\.\.HEAD|[Unreleased]: \1/compare/v${NEW_VERSION}...HEAD|" CHANGELOG.md

sed -i.bak "/^\[${CURRENT_VERSION}\]:/i\\
[${NEW_VERSION}]: https://github.com/abdullahmujahidali/Cellify/compare/v${CURRENT_VERSION}...v${NEW_VERSION}
" CHANGELOG.md

rm -f CHANGELOG.md.bak

echo -e "${GREEN}📝 Committing changes...${NC}"
git add package.json CHANGELOG.md
git commit -m "chore: release v${NEW_VERSION}"

echo -e "${GREEN}🏷️  Creating tag v${NEW_VERSION}...${NC}"
git tag -a "v${NEW_VERSION}" -m "Release v${NEW_VERSION}"

echo -e "${GREEN}📤 Pushing to remote...${NC}"
git push origin main
git push origin "v${NEW_VERSION}"

echo -e "${GREEN}📦 Publishing to npm...${NC}"
npm publish

echo -e "${GREEN}✅ Successfully released v${NEW_VERSION}!${NC}"
echo ""
echo -e "${YELLOW}Next steps:${NC}"
echo "1. Create GitHub release at: https://github.com/abdullahmujahidali/Cellify/releases/new?tag=v${NEW_VERSION}"
echo "2. Copy relevant CHANGELOG entries to the release notes"
