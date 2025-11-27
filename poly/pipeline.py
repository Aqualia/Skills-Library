from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Callable, Dict, Iterable, Optional

from .market_fetcher import MarketFetcher, MarketRecord


@dataclass
class PipelineResult:
    markets: Dict[str, MarketRecord]
    last_updated: Dict[str, datetime]
    step_timestamps: Dict[str, datetime]


class MarketPipeline:
    """Run fetch + processing steps while tracking freshness metadata."""

    def __init__(
        self,
        fetcher: MarketFetcher,
        processors: Optional[Iterable[Callable[[MarketRecord], None]]] = None,
        *,
        clock: Callable[[], datetime] | None = None,
    ) -> None:
        self.fetcher = fetcher
        self.processors = list(processors or [])
        self.clock: Callable[[], datetime] = clock or datetime.utcnow

    def run(
        self,
        *,
        existing: Optional[Dict[str, MarketRecord]] = None,
        force_refresh: bool = False,
        stale_after: Optional[timedelta] = None,
        limit: Optional[int] = None,
    ) -> PipelineResult:
        step_timestamps: Dict[str, datetime] = {"pipeline_started": self.clock()}
        markets = self.fetcher.fetch_markets(
            existing=existing,
            force_refresh=force_refresh,
            stale_after=stale_after,
            limit=limit,
        )
        step_timestamps["fetch_completed"] = self.clock()

        for processor in self.processors:
            label = getattr(processor, "__name__", "processor")
            for record in markets.values():
                processor(record)
                timestamp = self.clock()
                record.step_timestamps[label] = timestamp
                record.last_updated = max(record.last_updated, timestamp)
            step_timestamps[f"{label}_completed"] = self.clock()

        pipeline_completed = self.clock()
        step_timestamps["pipeline_completed"] = pipeline_completed
        last_updated = {market_id: record.last_updated for market_id, record in markets.items()}

        return PipelineResult(markets=markets, last_updated=last_updated, step_timestamps=step_timestamps)
