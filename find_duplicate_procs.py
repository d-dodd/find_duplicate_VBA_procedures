#!/usr/bin/env python3
"""
find_duplicate_procs.py

Scan a VBA_Export folder (the text export of an .xlsm project: .bas, .cls, .frm)
and report duplicate procedure names. Commented-out declarations are ignored.

Usage:
    python find_duplicate_procs.py "C:\\path\\to\\VBA_Export"
    python find_duplicate_procs.py "C:\\path\\to\\VBA_Export" -o dupes.md
    python find_duplicate_procs.py "C:\\path\\to\\VBA_Export" --ext .bas .cls .frm .txt

Behavior:
    - If duplicates are found, a Markdown report is written and the path is printed.
    - If no duplicates are found, nothing is written; the terminal shows
      "No duplicates found".

Exit codes:
    0 = no duplicates
    1 = duplicates found (report written)
    2 = usage / IO error

What counts as a procedure:
    Sub, Function, Property Get/Let/Set, and Declare (API) statements, with or
    without Public / Private / Friend / Static / Global modifiers, including
    declarations broken across lines with the "_" continuation character.

What counts as a duplicate:
    1. Same name twice in the same module          -> compile error in VBA
    2. Same public name in two standard (.bas)     -> "Ambiguous name" at
       modules                                        compile/run time
    3. Same name in two modules where at least one -> legal in VBA, but usually
       is a class/form module, or where the           copy-paste duplication
       procedures are Private/Friend

    Property Get/Let/Set sharing one name is NOT a duplicate (that is the normal
    property pattern). Two Property Gets with the same name in one module is.
"""

import argparse
import os
import re
import sys
from collections import defaultdict
from datetime import datetime

DEFAULT_EXTS = (".bas", ".cls", ".frm")
ENCODINGS = ("utf-8-sig", "cp1252", "latin-1")

PROC_RE = re.compile(
    r"""
    ^\s*
    (?P<modifiers>(?:(?:Public|Private|Friend|Global|Static)\s+)*)
    (?P<declare>Declare\s+(?:PtrSafe\s+)?)?
    (?P<kind>Sub|Function|Property\s+(?:Get|Let|Set))
    \s+
    (?P<name>\[[^\]]+\]|[^\W\d]\w*)
    \s*(?:\(|$|\s|:)
    """,
    re.IGNORECASE | re.VERBOSE,
)

REM_RE = re.compile(r"^rem\b\s*", re.IGNORECASE)
CC_RE = re.compile(r"^#\s*(If|ElseIf|Else|End\s+If)\b", re.IGNORECASE)
ATTR_NAME_RE = re.compile(r'^\s*Attribute\s+VB_Name\s*=\s*"([^"]*)"', re.IGNORECASE)
OPTION_PRIVATE_RE = re.compile(r"^\s*Option\s+Private\s+Module\b", re.IGNORECASE)

MODULE_TYPES = {
    ".bas": "Standard module",
    ".cls": "Class module",
    ".frm": "Form module",
}


class Occurrence(object):
    """One procedure declaration found in one module."""

    def __init__(self, name, kind, scope, module, module_type, rel_path,
                 line_no, text, is_declare, option_private, cc_path=()):
        self.name = name
        self.kind = kind                  # "Sub", "Function", "Property Get", ...
        self.scope = scope                # "Public", "Private", "Friend"
        self.module = module
        self.module_type = module_type
        self.rel_path = rel_path
        self.line_no = line_no
        self.text = text
        self.is_declare = is_declare
        self.option_private = option_private
        self.cc_path = cc_path        # (#If block id, branch index) per nesting level

    @property
    def sig_key(self):
        """Property accessors of different kinds may share a name; nothing else may."""
        k = self.kind.lower()
        if k.startswith("property"):
            return k
        return "proc"

    @property
    def display_kind(self):
        return ("Declare " + self.kind) if self.is_declare else self.kind


def read_text(path):
    """Read a VBA export file, trying the encodings Excel actually writes."""
    last_err = None
    for enc in ENCODINGS:
        try:
            with open(path, "r", encoding=enc, newline="") as fh:
                return fh.read()
        except UnicodeDecodeError as err:
            last_err = err
    raise last_err


def split_comment(line):
    """
    Split a physical line into (code, comment).

    An apostrophe inside a string literal is not a comment marker. VBA escapes a
    quote inside a string by doubling it, which the toggle handles for free.
    Returns comment=None when the line has no comment.
    """
    in_string = False
    for i, ch in enumerate(line):
        if ch == '"':
            in_string = not in_string
        elif ch == "'" and not in_string:
            return line[:i], line[i + 1:]
    return line, None


def logical_lines(text):
    """
    Yield (line_no, statement, is_commented) for each logical line.

    Physical lines joined by a trailing "_" are merged into one statement, and
    the line number reported is where the statement starts. Comment-only lines
    are yielded separately with is_commented=True so the caller can tell
    commented-out declarations apart from live ones.
    """
    pending = ""
    pending_start = None

    for line_no, raw in enumerate(text.splitlines(), 1):
        code, comment = split_comment(raw)

        # A "Rem" statement is a comment even though it has no apostrophe.
        if REM_RE.match(code.strip()):
            comment = REM_RE.sub("", code.strip(), count=1)
            code = ""

        stripped = code.strip()

        if not pending:
            if not stripped:
                if comment is not None and comment.strip():
                    yield (line_no, comment.strip().lstrip("'").strip(), True)
                continue
            pending_start = line_no
            pending = stripped
        else:
            pending = pending + " " + stripped

        if pending.rstrip().endswith("_"):
            pending = pending.rstrip()[:-1]
            continue

        yield (pending_start, pending, False)
        pending = ""

    if pending:
        yield (pending_start, pending, False)


def parse_declaration(statement):
    """Return (name, kind, scope, is_declare) if the statement declares a procedure."""
    match = PROC_RE.match(statement)
    if not match:
        return None

    modifiers = (match.group("modifiers") or "").lower()
    if "private" in modifiers:
        scope = "Private"
    elif "friend" in modifiers:
        scope = "Friend"
    else:
        scope = "Public"          # VBA procedures are Public by default

    kind = re.sub(r"\s+", " ", match.group("kind")).title()

    name = match.group("name").strip()
    if name.startswith("["):
        name = name[1:-1].strip()

    return name, kind, scope, bool(match.group("declare"))


def scan_file(path, root):
    """Parse one export file. Returns (occurrences, commented_out_count)."""
    text = read_text(path)
    rel_path = os.path.relpath(path, root)
    ext = os.path.splitext(path)[1].lower()
    module_type = MODULE_TYPES.get(ext, "Other")
    module = os.path.splitext(os.path.basename(path))[0]
    option_private = False

    occurrences = []
    commented = 0
    header_done = False
    cc_stack = []
    cc_blocks = 0

    for line_no, statement, is_commented in logical_lines(text):
        if not is_commented:
            cc = CC_RE.match(statement)
            if cc:
                directive = re.sub(r"\s+", " ", cc.group(1)).lower()
                if directive == "if":
                    cc_blocks += 1
                    cc_stack.append([(rel_path, cc_blocks), 0])
                elif directive in ("elseif", "else"):
                    if cc_stack:
                        cc_stack[-1][1] += 1
                elif cc_stack:
                    cc_stack.pop()
                continue

        if not is_commented and not header_done:
            attr = ATTR_NAME_RE.match(statement)
            if attr and attr.group(1).strip():
                module = attr.group(1).strip()
            if OPTION_PRIVATE_RE.match(statement):
                option_private = True

        parsed = parse_declaration(statement)
        if not parsed:
            continue

        if is_commented:
            commented += 1
            continue

        header_done = True
        name, kind, scope, is_declare = parsed
        occurrences.append(
            Occurrence(
                name=name,
                kind=kind,
                scope=scope,
                module=module,
                module_type=module_type,
                rel_path=rel_path,
                line_no=line_no,
                text=statement,
                is_declare=is_declare,
                option_private=option_private,
                cc_path=tuple((bid, branch) for bid, branch in cc_stack),
            )
        )

    # Attribute VB_Name may appear after the first declaration in a .frm header,
    # so backfill the module name onto everything found in this file.
    for occ in occurrences:
        occ.module = module
        occ.option_private = option_private

    return occurrences, commented


def collect_files(root, exts):
    found = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames.sort()
        for filename in sorted(filenames):
            if os.path.splitext(filename)[1].lower() in exts:
                found.append(os.path.join(dirpath, filename))
    return found


def mutually_exclusive(first, second):
    """
    True when two declarations sit in different branches of the same
    #If / #Else block, so only one of them is ever compiled. This is the
    standard 32-bit / 64-bit API Declare pattern and is not a duplicate.
    """
    for left, right in zip(first.cc_path, second.cc_path):
        if left[0] == right[0] and left[1] != right[1]:
            return True
    return False


def collides(first, second):
    """True when two same-named declarations genuinely clash."""
    if mutually_exclusive(first, second):
        return False
    # Property Get/Let/Set may share one name; that is the normal pattern.
    if first.sig_key != "proc" and second.sig_key != "proc":
        return first.sig_key == second.sig_key
    return True


def conflicting_subset(group):
    """Return only the declarations from a same-named group that actually clash."""
    keep = []
    for i, first in enumerate(group):
        for j, second in enumerate(group):
            if i != j and collides(first, second):
                keep.append(first)
                break
    return keep


def analyze(occurrences):
    """Bucket duplicates into same-module, ambiguous-public, and cross-module."""
    by_name = defaultdict(list)
    for occ in occurrences:
        by_name[occ.name.lower()].append(occ)

    same_module = []      # (name, [occ, ...])
    ambiguous = []
    cross_module = []

    for key in sorted(by_name):
        group = by_name[key]
        if len(group) < 2:
            continue

        by_module = defaultdict(list)
        for occ in group:
            by_module[occ.module].append(occ)

        for module in sorted(by_module):
            in_module = sorted(by_module[module], key=lambda o: o.line_no)
            clashing = conflicting_subset(in_module)
            if len(clashing) > 1:
                same_module.append((clashing[0].name, clashing))

        if len(by_module) < 2:
            continue

        # One representative per module decides whether the modules collide.
        reps = [sorted(by_module[m], key=lambda o: o.line_no)[0]
                for m in sorted(by_module)]
        clashing_modules = set(occ.module for occ in conflicting_subset(reps))
        if len(clashing_modules) < 2:
            continue

        involved = [occ for occ in group if occ.module in clashing_modules]
        ordered = sorted(involved, key=lambda o: (o.module.lower(), o.line_no))

        public_bas_modules = set(
            occ.module for occ in involved
            if occ.scope == "Public" and occ.module_type == "Standard module")

        if len(public_bas_modules) > 1:
            ambiguous.append((ordered[0].name, ordered))
        else:
            cross_module.append((ordered[0].name, ordered))

    return same_module, ambiguous, cross_module


def md_escape(value):
    return str(value).replace("|", "\\|")


def render_table(rows):
    lines = [
        "| Module | Module type | File | Line | Kind | Scope | Declaration |",
        "| --- | --- | --- | --- | --- | --- | --- |",
    ]
    for occ in rows:
        lines.append("| {0} | {1} | {2} | {3} | {4} | {5} | `{6}` |".format(
            md_escape(occ.module),
            md_escape(occ.module_type),
            md_escape(occ.rel_path),
            occ.line_no,
            md_escape(occ.display_kind),
            md_escape(occ.scope + (" (Option Private Module)" if occ.option_private else "")),
            md_escape(occ.text.replace("`", "'")),
        ))
    return lines


def build_report(root, file_count, proc_count, commented_count,
                 same_module, ambiguous, cross_module):
    out = []
    out.append("# Duplicate VBA Procedures")
    out.append("")
    out.append("Source folder: `{0}`".format(root))
    out.append("")
    out.append("Generated: {0}".format(datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    out.append("")
    out.append("- Files scanned: {0}".format(file_count))
    out.append("- Live procedure declarations: {0}".format(proc_count))
    out.append("- Commented-out declarations ignored: {0}".format(commented_count))
    out.append("- Duplicates in the same module: {0}".format(len(same_module)))
    out.append("- Ambiguous public names across standard modules: {0}".format(len(ambiguous)))
    out.append("- Repeated names across modules (legal): {0}".format(len(cross_module)))
    out.append("")
    out.append("A name can appear in more than one section below: a routine declared "
               "twice in one module and again in a second module is both a compile "
               "error and an ambiguity.")
    out.append("")

    sections = [
        (
            "Duplicates in the same module",
            "Two declarations of the same name in one module. VBA will not "
            "compile this; one of them has to go.",
            same_module,
        ),
        (
            "Ambiguous public names across standard modules",
            "The same public name in more than one standard (.bas) module. VBA "
            "raises \"Ambiguous name detected\" unless every call site is "
            "qualified with the module name. Make all but one Private, or rename.",
            ambiguous,
        ),
        (
            "Repeated names across modules",
            "Legal in VBA because the names are scoped to a class, form, or to a "
            "single module via Private/Friend. Worth reviewing anyway: these are "
            "usually copy-pasted routines that have since drifted apart.",
            cross_module,
        ),
    ]

    for title, blurb, items in sections:
        out.append("## {0}".format(title))
        out.append("")
        if not items:
            out.append("None found.")
            out.append("")
            continue
        out.append(blurb)
        out.append("")
        for name, group in items:
            out.append("### `{0}` ({1} declarations)".format(name, len(group)))
            out.append("")
            out.extend(render_table(group))
            out.append("")

    return "\n".join(out) + "\n"


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Find duplicate procedures in a VBA_Export folder.")
    parser.add_argument("folder", help="Path to the VBA_Export folder")
    parser.add_argument("-o", "--output", default=None,
                        help="Report path (default: duplicate_procedures_<timestamp>.md "
                             "in the current directory)")
    parser.add_argument("--ext", nargs="+", default=list(DEFAULT_EXTS),
                        help="File extensions to scan (default: .bas .cls .frm)")
    args = parser.parse_args(argv)

    root = os.path.abspath(args.folder)
    if not os.path.isdir(root):
        print("Not a folder: {0}".format(root), file=sys.stderr)
        return 2

    exts = set(e.lower() if e.startswith(".") else "." + e.lower() for e in args.ext)
    files = collect_files(root, exts)
    if not files:
        print("No {0} files found under {1}".format("/".join(sorted(exts)), root),
              file=sys.stderr)
        return 2

    occurrences = []
    commented_count = 0
    for path in files:
        try:
            found, commented = scan_file(path, root)
        except (OSError, UnicodeDecodeError) as err:
            print("Skipped {0}: {1}".format(path, err), file=sys.stderr)
            continue
        occurrences.extend(found)
        commented_count += commented

    same_module, ambiguous, cross_module = analyze(occurrences)

    if not (same_module or ambiguous or cross_module):
        print("No duplicates found")
        return 0

    if args.output:
        out_path = os.path.abspath(args.output)
    else:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = os.path.abspath("duplicate_procedures_{0}.md".format(stamp))

    report = build_report(root, len(files), len(occurrences), commented_count,
                          same_module, ambiguous, cross_module)

    out_dir = os.path.dirname(out_path)
    if out_dir and not os.path.isdir(out_dir):
        os.makedirs(out_dir)
    with open(out_path, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(report)

    print("Duplicates found: {0} in-module, {1} ambiguous public, {2} cross-module".format(
        len(same_module), len(ambiguous), len(cross_module)))
    print("Report: {0}".format(out_path))
    return 1


if __name__ == "__main__":
    sys.exit(main())