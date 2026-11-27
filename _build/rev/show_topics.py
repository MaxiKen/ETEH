import json, sys
inv=json.load(open("/home/user/ETEH/_build/_stage/inventory.json"))
vol, mods = sys.argv[1], sys.argv[2].split(",")
width=int(sys.argv[3]) if len(sys.argv)>3 else 260
for t in inv[vol]:
    if t["module"] in mods:
        body=" ¶ ".join(x for x in t["texts"] if not x.isupper() and not x.startswith(("WORKED EXAMPLE","fx ")) )
        print(f"## {t['tag']} {t['title']}  [{len(t['slides'])} sl]")
        print("   ", body[:width])
        if t["tables"]: print("    TBL:", " | ".join(t["tables"][0][0]))
