"""The timestamp stamped into generated documentation.

`marinade document` embeds a generation time in `documentation.md` and
`model_spec.json`, so running it twice over one IR database produced
byte-different artifacts — awkward for a tool whose outputs are otherwise
deterministic, and noisy for anything that diffs or hashes them (a docs CI
check, or a user comparing two runs).

`SOURCE_DATE_EPOCH` is the reproducible-builds convention for exactly this:
when set, it pins the timestamp and the output becomes byte-stable; when
unset, the wall clock is used and humans keep a meaningful "Generated:" line.
"""

import os
from datetime import UTC, datetime


def generation_timestamp() -> str:
    """ISO-8601 UTC timestamp, pinned by SOURCE_DATE_EPOCH when set."""
    source_date_epoch = os.environ.get("SOURCE_DATE_EPOCH")
    if source_date_epoch:
        try:
            moment = datetime.fromtimestamp(int(source_date_epoch), tz=UTC)
        except (ValueError, OverflowError, OSError):
            # A malformed value must not fail a documentation run; fall back to
            # the wall clock, which is the behaviour without the variable set.
            moment = datetime.now(UTC)
    else:
        moment = datetime.now(UTC)
    return moment.isoformat().replace("+00:00", "Z")
