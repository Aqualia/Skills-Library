from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import Callable, Dict, Iterable, Optional, Protocol, Tuple


class MarketClient(Protocol):
    """Client that can list and hydrate market records."""

    def fetch_page(self, *, cursor: Optional[str] = None, limit: int = 0) -> Tuple[Iterable[dict], Optional[str]]:
        """Return a page of market summaries and the next cursor if pagination continues."""

    def fetch_detail(self, market_id: str) -> dict:
        """Return the fully hydrated market payload for a single market."""


@dataclass
class MarketRecord:
    """Represents the hydrated market payload and freshness metadata."""

    market_id: str
    payload: dict
    last_updated: datetime
    step_timestamps: Dict[str, datetime] = field(default_factory=dict)


class MarketFetcher:
    """Fetch markets while respecting freshness requirements."""

    def __init__(
        self,
        client: MarketClient,
        *,
        default_limit: int = 200,
        stale_after: timedelta = timedelta(hours=1),
        clock: Callable[[], datetime] | None = None,
    ) -> None:
        self.client = client
        self.default_limit = default_limit
        self.stale_after = stale_after
        self.clock: Callable[[], datetime] = clock or datetime.utcnow

    def _is_stale(self, record: Optional[MarketRecord], stale_after: Optional[timedelta]) -> bool:
        if record is None:
            return True
        threshold = stale_after or self.stale_after
        return self.clock() - record.last_updated >= threshold

    def fetch_markets(
        self,
        *,
        existing: Optional[Dict[str, MarketRecord]] = None,
        force_refresh: bool = False,
        stale_after: Optional[timedelta] = None,
        limit: Optional[int] = None,
    ) -> Dict[str, MarketRecord]:
        """Fetch markets while skipping fresh entries unless forced.

        Args:
            existing: Previously fetched markets keyed by id.
            force_refresh: If True, always refetch even when recent.
            stale_after: Optional override for freshness threshold.
            limit: Optional override for page size when listing markets.
        """

        records: Dict[str, MarketRecord] = {}
        cursor: Optional[str] = None
        page_limit = limit or self.default_limit

        while True:
            summaries, cursor = self.client.fetch_page(cursor=cursor, limit=page_limit)
            for summary in summaries:
                market_id = summary["id"]
                record = existing.get(market_id) if existing else None
                needs_refresh = force_refresh or self._is_stale(record, stale_after)
                if needs_refresh:
                    payload = self.client.fetch_detail(market_id)
                    record = MarketRecord(
                        market_id=market_id,
                        payload=payload,
                        last_updated=self.clock(),
                        step_timestamps={},
                    )
                records[market_id] = record
            if cursor is None:
                break

        return records
