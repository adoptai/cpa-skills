#!/bin/bash
#
# Validate every skill against the Agent Skills specification, plus checks
# specific to this repository.
#
#   https://agentskills.io/specification.md
#
# Spec:      name (matches dir, 1-64 chars, lowercase/digits/hyphens),
#            description (1-1024 chars), SKILL.md under 500 lines,
#            optional dirs limited to references/ scripts/ assets/
#
# This repo: every skill listed in .claude-plugin/marketplace.json,
#            every referenced script exists, Python compiles,
#            and nothing that looks like real client data is committed.
#
# Usage:  ./validate-skills.sh
# Exit:   0 clean, 1 errors found

set -uo pipefail

RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[1;33m'; NC='\033[0m'

SKILLS_DIR="skills"
MANIFEST=".claude-plugin/marketplace.json"
SPEC_DIRS=("references" "scripts" "assets")

ERRORS=0; WARNINGS=0; PASSED=0

echo "Validating skills against the Agent Skills specification"
echo "========================================================"
echo "  https://agentskills.io/specification.md"
echo ""

if [[ ! -d "$SKILLS_DIR" ]]; then
    echo -e "${RED}No $SKILLS_DIR/ directory. Run this from the repository root.${NC}"
    exit 1
fi

# ------------------------------------------------------------------ per skill
for skill_dir in "$SKILLS_DIR"/*/; do
    name=$(basename "$skill_dir")
    file="$skill_dir/SKILL.md"
    errs=(); warns=()

    if [[ ! -f "$file" ]]; then
        echo -e "${RED}FAIL  $name${NC}"
        echo "        no SKILL.md"
        ((ERRORS++)); continue
    fi

    frontmatter=$(awk '/^---$/{c++; next} c==1' "$file")
    if [[ -z "$frontmatter" ]]; then
        echo -e "${RED}FAIL  $name${NC}"
        echo "        no YAML frontmatter"
        ((ERRORS++)); continue
    fi

    # ---- name
    fm_name=$(echo "$frontmatter" | grep "^name:" | head -1 | sed 's/^name: *//' | tr -d ' ')
    if [[ -z "$fm_name" ]]; then
        errs+=("missing 'name' in frontmatter")
    elif [[ "$fm_name" != "$name" ]]; then
        errs+=("name mismatch: directory '$name' vs frontmatter '$fm_name'")
    elif ! [[ "$fm_name" =~ ^[a-z0-9]([a-z0-9-]{0,62}[a-z0-9])?$ ]]; then
        errs+=("invalid name '$fm_name' (lowercase, digits, hyphens; no leading/trailing hyphen)")
    elif [[ "$fm_name" == *"--"* ]]; then
        errs+=("name '$fm_name' contains consecutive hyphens")
    fi

    # ---- description
    desc=$(echo "$frontmatter" | grep "^description:" | head -1 | sed 's/^description: *//')
    desc="${desc%\"}"; desc="${desc#\"}"
    if [[ -z "$desc" ]]; then
        errs+=("missing 'description' in frontmatter")
    else
        len=${#desc}
        if (( len > 1024 )); then
            errs+=("description is $len chars (spec limit 1024)")
        elif (( len < 80 )); then
            warns+=("description is only $len chars - likely too thin to trigger reliably")
        fi
        if ! echo "$desc" | grep -qi "when\|use this\|mention"; then
            warns+=("description lacks trigger phrasing ('when', 'use this', 'mentions')")
        fi
    fi

    # ---- body size
    lines=$(wc -l < "$file")
    if (( lines >= 500 )); then
        errs+=("SKILL.md is $lines lines (spec: under 500 - move detail into references/)")
    elif (( lines > 400 )); then
        warns+=("SKILL.md is $lines lines - approaching the 500-line limit")
    fi

    # ---- optional directories must be spec-named
    while IFS= read -r sub; do
        [[ -z "$sub" ]] && continue
        base=$(basename "$sub")
        ok=0
        for allowed in "${SPEC_DIRS[@]}"; do
            [[ "$base" == "$allowed" ]] && ok=1
        done
        if (( ! ok )); then
            errs+=("non-spec directory '$base/' (allowed: ${SPEC_DIRS[*]}) - will not be bundled")
        fi
    done < <(find "$skill_dir" -mindepth 1 -maxdepth 1 -type d 2>/dev/null)

    # ---- scripts referenced in SKILL.md actually exist
    # Capture any leading path so a deliberate cross-skill reference
    # (../other-skill/scripts/foo.py) can be recognised and skipped - those are
    # legitimate when documenting a chain between skills.
    while IFS= read -r ref; do
        [[ -z "$ref" ]] && continue
        [[ "$ref" == *"../"* ]] && continue      # cross-skill reference, not local
        ref="${ref#./}"
        if [[ ! -f "$skill_dir/$ref" ]]; then
            errs+=("SKILL.md references '$ref' which does not exist")
        fi
    done < <(grep -oE '[A-Za-z0-9_./-]*(scripts|references|assets)/[A-Za-z0-9_./-]+' \
             "$file" | sort -u)

    # ---- Python compiles
    while IFS= read -r py; do
        [[ -z "$py" ]] && continue
        if ! python3 -m py_compile "$py" 2>/dev/null; then
            errs+=("$(basename "$py") does not compile")
        fi
    done < <(find "$skill_dir" -name "*.py" 2>/dev/null)

    # ---- listed in the plugin manifest
    if [[ -f "$MANIFEST" ]] && ! grep -q "\./skills/$name\"" "$MANIFEST"; then
        errs+=("not listed in $MANIFEST - it will not install via the plugin marketplace")
    fi

    # ---- report
    if (( ${#errs[@]} )); then
        echo -e "${RED}FAIL  $name${NC}"
        for e in "${errs[@]}"; do echo -e "        ${RED}error:${NC}   $e"; done
        for w in "${warns[@]:-}"; do [[ -n "$w" ]] && echo -e "        ${YELLOW}warning:${NC} $w"; done
        ((ERRORS++))
    elif (( ${#warns[@]} )); then
        echo -e "${YELLOW}WARN  $name${NC}  (desc ${#desc} chars, $lines lines)"
        for w in "${warns[@]}"; do echo -e "        ${YELLOW}warning:${NC} $w"; done
        ((WARNINGS++))
    else
        echo -e "${GREEN}ok    $name${NC}  (desc ${#desc} chars, $lines lines)"
        ((PASSED++))
    fi
done

find . -name "__pycache__" -type d -exec rm -rf {} + 2>/dev/null

# ------------------------------------------------------------------ repo-wide
echo ""
echo "Repository checks"
echo "-----------------"

for f in README.md LICENSE CONTRIBUTING.md AGENTS.md CLAUDE.md VERSIONS.md \
         .gitignore "$MANIFEST" .claude-plugin/plugin.json; do
    if [[ -f "$f" ]]; then
        echo -e "${GREEN}ok    ${NC}$f"
    else
        echo -e "${RED}FAIL  ${NC}$f missing"
        ((ERRORS++))
    fi
done

for j in "$MANIFEST" .claude-plugin/plugin.json; do
    if [[ -f "$j" ]]; then
        if python3 -c "import json,sys; json.load(open('$j'))" 2>/dev/null; then
            echo -e "${GREEN}ok    ${NC}$j is valid JSON"
        else
            echo -e "${RED}FAIL  ${NC}$j is not valid JSON"
            ((ERRORS++))
        fi
    fi
done

# manifest must not list a skill that no longer exists
if [[ -f "$MANIFEST" ]]; then
    while IFS= read -r listed; do
        [[ -z "$listed" ]] && continue
        if [[ ! -d "$listed" ]]; then
            echo -e "${RED}FAIL  ${NC}$MANIFEST lists '$listed' which does not exist"
            ((ERRORS++))
        fi
    done < <(grep -oE '\./skills/[a-z0-9-]+' "$MANIFEST" | sed 's|^\./||' | sort -u)
fi

# Client data must never be committed. These skills read bank statements, GLs,
# and tax returns, so a stray fixture is a genuine confidentiality incident.
echo ""
echo "Client data check"
echo "-----------------"
leaks=0
while IFS= read -r f; do
    case "$f" in
        ./.git/*|./dist/*|./markdown/*) continue ;;
    esac
    echo -e "${RED}FAIL  ${NC}$f should not be committed (spreadsheet/PDF/export)"
    ((leaks++))
done < <(find . -type f \( -name "*.xlsx" -o -name "*.xls" -o -name "*.pdf" \
         -o -name "*.qbo" -o -name "*.qbw" -o -name "*.ofx" \) 2>/dev/null)

# SSN / EIN shaped strings in tracked text files
while IFS= read -r hit; do
    echo -e "${RED}FAIL  ${NC}possible SSN/EIN pattern: $hit"
    ((leaks++))
done < <(grep -rInE '\b[0-9]{3}-[0-9]{2}-[0-9]{4}\b|\b[0-9]{2}-[0-9]{7}\b' \
         --include="*.md" --include="*.py" --include="*.csv" --include="*.json" \
         . 2>/dev/null | grep -v "^./.git/" | head -10)

if (( leaks == 0 )); then
    echo -e "${GREEN}ok    ${NC}no spreadsheets, PDFs, exports, or TIN-shaped strings committed"
else
    ((ERRORS += leaks))
fi

# ------------------------------------------------------------------ summary
echo ""
echo "========================================================"
echo -e "  ${GREEN}passed:   $PASSED${NC}"
(( WARNINGS )) && echo -e "  ${YELLOW}warnings: $WARNINGS${NC}"
(( ERRORS ))   && echo -e "  ${RED}errors:   $ERRORS${NC}"
echo ""

if (( ERRORS == 0 )); then
    echo -e "${GREEN}All skills valid.${NC}"
    exit 0
fi
echo -e "${RED}$ERRORS error(s) need fixing.${NC}"
exit 1
