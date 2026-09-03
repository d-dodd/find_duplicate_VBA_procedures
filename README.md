# find_VBA_duplicates
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
