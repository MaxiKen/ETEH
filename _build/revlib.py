"""Helper for topic copy files in _build/rev/ (see rev/DESIGN.md)."""


def T(tag, **kw):
    kw["tag"] = tag
    for k in ("what", "like", "tip", "trap", "case", "steps", "so", "who"):
        assert k in kw, f"{tag}: missing {k}"
    return kw
