"""Check that functions.csv records what each function used to be called.

Usage: python tools/verify_previous_names.py [functions.csv] [baseline]

v2.0.0 renamed all 130 functions from six module prefixes to one, which broke every
formula written against the old names. `functions.csv` carries a `previous_name` column
so a reader can look up the replacement, and `tools/released-names-v1.2.6.txt` records
the 130 names the last release before the rename actually shipped.

The build fills the column and asserts as it goes, but the build is not what people read:
the committed CSV is. This checks the published file against the published baseline, so a
stale or hand-edited CSV fails here rather than misinforming somebody mid-migration.

Four things must hold:

  * every name the baseline records is claimed by exactly one function, since an
    unclaimed one is a function that disappeared without a forwarding address
  * no `previous_name` names something the baseline does not, since a predecessor that
    never shipped is worse than no predecessor at all
  * a function added since the baseline records nothing, which is honest, rather than a
    plausible-looking name derived from the build's own intermediate naming
  * each function's claimed predecessor is recognisably its own, not merely some unused
    baseline name. The first three checks together prove the map is a bijection, which is
    not the same as proving it is the right one: swapping two unrelated functions'
    predecessors satisfies all three and still sends a reader to the wrong function. The
    rename only ever appended to a bare name, adding a `B`, `E` or `U` tag or a module
    word, so the new bare name must begin with the old one.

A blank `previous_name` is therefore expected and allowed. It means "new since v1.2.6".
"""
import csv
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Every function name carries a λ, and a Windows console defaults to cp1252, which
# cannot encode it: without this the check dies printing its own result.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

CSV = sys.argv[1] if len(sys.argv) > 1 else "functions.csv"
BASELINE = (sys.argv[2] if len(sys.argv) > 2
            else os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              "released-names-v1.2.6.txt"))
FLOOR = 100          # the library ships 130 functions; reading far fewer is a bug


def main():
    with open(BASELINE, encoding="utf-8") as fh:
        released = {ln.strip() for ln in fh if ln.strip() and not ln.startswith("#")}
    with open(CSV, encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.DictReader(fh))

    if not rows:
        sys.exit("no rows read from %s" % CSV)
    if "previous_name" not in rows[0]:
        sys.exit("%s has no previous_name column" % CSV)
    if len(rows) < FLOOR or len(released) < FLOOR:
        sys.exit("only %d functions and %d baseline names read: the check is not reading them"
                 % (len(rows), len(released)))

    failures = []
    claimed = {}
    for row in rows:
        was = row["previous_name"].strip()
        if not was:
            continue                     # added since the baseline, correctly recording nothing
        if was not in released:
            failures.append("%s claims to replace %s, which %s never shipped"
                            % (row["function"], was, os.path.basename(BASELINE)))
        elif was in claimed:
            failures.append("%s is claimed by both %s and %s"
                            % (was, claimed[was], row["function"]))
        else:
            claimed[was] = row["function"]
            # the rename only ever appended to a bare name, so an unrelated pairing shows
            # up here even though it satisfies every count
            old_bare = was.rsplit(".", 1)[1].rstrip("λ")
            new_bare = row["function"].split(".", 1)[1]
            if not new_bare.startswith(old_bare):
                failures.append("%s is not a renaming of %s: %s does not begin with %s"
                                % (row["function"], was, new_bare, old_bare))

    # The index is read by people and by machines, so the two columns nothing else checks
    # get checked here. Each has shipped wrong: nb.Depreciateλ published a signature cut off
    # mid-parameter-list because its help wraps onto a second row, and 31 descriptions
    # carried the raw OOXML escape for a line break out of a Name Manager comment.
    for row in rows:
        sig = row.get("signature", "")
        blurb = row.get("description", "")
        if sig.count("(") != sig.count(")"):
            failures.append("%s publishes a signature with unbalanced brackets: %s"
                            % (row["function"], sig))
        for field, text in (("signature", sig), ("description", blurb)):
            if "_x00" in text:
                failures.append("%s's %s carries a raw OOXML escape: %s"
                                % (row["function"], field, text[:60]))
        if blurb.startswith("→"):
            failures.append("%s's description starts with the help table's column delimiter"
                            % row["function"])
        if not sig or not blurb:
            failures.append("%s publishes an empty %s"
                            % (row["function"], "signature" if not sig else "description"))

    for was in sorted(released - set(claimed)):
        failures.append("%s shipped in the baseline and no function claims it" % was)

    if failures:
        print("FAIL: %d problem(s) in %s" % (len(failures), CSV))
        for item in failures:
            print("  - %s" % item)
        return 1
    fresh = sum(1 for row in rows if not row["previous_name"].strip())
    print("OK: %d functions, %d of them replacing a name from %s, %d new since it"
          % (len(rows), len(claimed), os.path.basename(BASELINE), fresh))
    return 0


if __name__ == "__main__":
    sys.exit(main())
