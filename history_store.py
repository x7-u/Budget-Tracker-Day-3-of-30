"""Day 3. Persistent history store for anomaly detection.

Every successful run appends one JSONL record per row keyed by
``(cost_centre, category, line_type)`` to ``outputs/history.jsonl``.
On the next run we compute a z-score per key against the history
when at least ``MIN_OBSERVATIONS`` prior records exist. Rows with an
absolute z-score over ``Z_THRESHOLD`` are flagged as statistical
anomalies, separate from the tolerance-band RAG.

The store is intentionally simple. It is a moving log; it does not
deduplicate or compress. The user can clear it via the API.
"""
from __future__ import annotations

import datetime as dt
import json
import math
import threading
import uuid
from dataclasses import dataclass
from pathlib import Path

DEFAULT_HISTORY_PATH = Path(__file__).resolve().parent / "outputs" / "history.jsonl"
MIN_OBSERVATIONS = 3
Z_THRESHOLD = 2.0
MAX_RECORDS_LOADED = 50_000


@dataclass
class AnomalyMark:
    z_score: float
    mean: float
    stdev: float
    n: int
    severity: str        # "high" if |z| >= 3, else "moderate"


class HistoryStore:
    def __init__(self, path: Path | None = None) -> None:
        self.path = Path(path) if path is not None else DEFAULT_HISTORY_PATH
        self._lock = threading.Lock()

    # ---- mutations ----------------------------------------------------

    def append_run(self, *, company: str, period_label: str, rows) -> int:
        """Append one record per row from a completed analysis. Returns count."""
        ts = dt.datetime.now(dt.UTC).replace(microsecond=0, tzinfo=None).isoformat() + "Z"
        run_id = uuid.uuid4().hex[:12]
        n = 0
        with self._lock:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            with self.path.open("a", encoding="utf-8") as f:
                for r in rows:
                    if r.variance_pct is None:
                        continue
                    record = {
                        "run_id": run_id,
                        "ts": ts,
                        "company": company,
                        "period_label": period_label,
                        "period": r.period,
                        "cost_centre": r.cost_centre,
                        "category": r.category,
                        "line_type": r.line_type,
                        "variance_pct": float(r.variance_pct),
                        "variance": float(r.variance),
                        "budget": float(r.budget),
                    }
                    f.write(json.dumps(record, ensure_ascii=False) + "\n")
                    n += 1
        return n

    def clear(self) -> int:
        with self._lock:
            entries = self._load_all()
            n = len(entries)
            if self.path.exists():
                self.path.write_text("", encoding="utf-8")
            return n

    # ---- queries ------------------------------------------------------

    def keys_with_enough_history(self, *, company: str | None = None) -> dict[tuple[str, str, str], list[float]]:
        """Return {(cost_centre, category, line_type): [variance_pct, ...]} buckets.

        Only buckets with at least MIN_OBSERVATIONS records are included.
        Filtered to a single company when supplied, since variance distributions
        are not portable between organisations.
        """
        buckets: dict[tuple[str, str, str], list[float]] = {}
        for rec in self._load_all():
            if company and rec.get("company") != company:
                continue
            key = (rec.get("cost_centre", ""), rec.get("category", ""), rec.get("line_type", ""))
            v = rec.get("variance_pct")
            if v is None:
                continue
            buckets.setdefault(key, []).append(float(v))
        return {k: vs for k, vs in buckets.items() if len(vs) >= MIN_OBSERVATIONS}

    def stats(self) -> dict:
        items = self._load_all()
        return {
            "records": len(items),
            "keys": len({(i.get("cost_centre"), i.get("category"), i.get("line_type")) for i in items}),
        }

    # ---- internals ----------------------------------------------------

    def _load_all(self) -> list[dict]:
        if not self.path.exists():
            return []
        try:
            text = self.path.read_text(encoding="utf-8")
        except OSError:
            return []
        out: list[dict] = []
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
            try:
                out.append(json.loads(line))
            except json.JSONDecodeError:
                continue
            if len(out) >= MAX_RECORDS_LOADED:
                break
        return out


# ---- Anomaly scoring --------------------------------------------------

def score_rows(rows, history: dict[tuple[str, str, str], list[float]]) -> int:
    """Mutate VarianceRows in-place: set ``z_score`` (and the mean / stdev /
    n that produced it) where history allows.

    Returns the number of rows flagged as anomalies (|z| >= Z_THRESHOLD).
    """
    flagged = 0
    for r in rows:
        if r.variance_pct is None:
            continue
        key = (r.cost_centre, r.category, r.line_type)
        bucket = history.get(key)
        if not bucket:
            continue
        mean, sd = _mean_stdev(bucket)
        if sd == 0:
            continue
        z = (r.variance_pct - mean) / sd
        r.z_score = z
        r.z_mean = mean
        r.z_stdev = sd
        r.z_n = len(bucket)
        if abs(z) >= Z_THRESHOLD:
            flagged += 1
    return flagged


def _mean_stdev(values: list[float]) -> tuple[float, float]:
    n = len(values)
    if n == 0:
        return 0.0, 0.0
    mean = sum(values) / n
    if n == 1:
        return mean, 0.0
    sq_sum = sum((v - mean) ** 2 for v in values)
    sd = math.sqrt(sq_sum / (n - 1))
    return mean, sd
