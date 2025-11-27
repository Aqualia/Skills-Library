from datetime import datetime, timedelta

from poly.market_fetcher import MarketFetcher, MarketRecord


class RecordingClient:
    def __init__(self, pages: list[list[dict]], payloads: dict[str, dict]):
        self.pages = pages
        self.payloads = payloads
        self.fetched_ids: list[str] = []
        self.last_limit = None

    def fetch_page(self, *, cursor=None, limit=0):
        self.last_limit = limit
        if cursor is None:
            page_index = 0
        else:
            page_index = int(cursor)
        try:
            page = self.pages[page_index]
        except IndexError:
            return [], None
        next_cursor = str(page_index + 1) if page_index + 1 < len(self.pages) else None
        return page, next_cursor

    def fetch_detail(self, market_id: str):
        self.fetched_ids.append(market_id)
        return self.payloads[market_id]


def frozen_clock(start: datetime):
    current = {"value": start}

    def _clock():
        return current["value"]

    def advance(seconds: int):
        current["value"] = current["value"] + timedelta(seconds=seconds)

    return _clock, advance


def test_skip_fresh_market_without_force():
    clock, advance = frozen_clock(datetime(2024, 1, 1, 12, 0, 0))
    client = RecordingClient(pages=[[{"id": "mkt-1"}]], payloads={"mkt-1": {"name": "alpha"}})
    fresh_record = MarketRecord("mkt-1", {"name": "alpha"}, last_updated=clock(), step_timestamps={})
    fetcher = MarketFetcher(client, stale_after=timedelta(minutes=30), clock=clock)

    result = fetcher.fetch_markets(existing={"mkt-1": fresh_record})

    assert result["mkt-1"].payload == {"name": "alpha"}
    assert client.fetched_ids == []


def test_refetch_when_stale():
    clock, advance = frozen_clock(datetime(2024, 1, 1, 12, 0, 0))
    client = RecordingClient(pages=[[{"id": "mkt-1"}]], payloads={"mkt-1": {"name": "alpha-new"}})
    stale_record = MarketRecord("mkt-1", {"name": "alpha"}, last_updated=clock(), step_timestamps={})
    advance(3600)
    fetcher = MarketFetcher(client, stale_after=timedelta(minutes=30), clock=clock)

    result = fetcher.fetch_markets(existing={"mkt-1": stale_record})

    assert result["mkt-1"].payload == {"name": "alpha-new"}
    assert client.fetched_ids == ["mkt-1"]


def test_force_refresh_overrides_freshness():
    clock, _ = frozen_clock(datetime(2024, 1, 1, 12, 0, 0))
    client = RecordingClient(pages=[[{"id": "mkt-1"}]], payloads={"mkt-1": {"name": "alpha-new"}})
    fresh_record = MarketRecord("mkt-1", {"name": "alpha"}, last_updated=clock(), step_timestamps={})
    fetcher = MarketFetcher(client, clock=clock)

    result = fetcher.fetch_markets(existing={"mkt-1": fresh_record}, force_refresh=True)

    assert result["mkt-1"].payload == {"name": "alpha-new"}
    assert client.fetched_ids == ["mkt-1"]


def test_pagination_uses_limit_override():
    clock, _ = frozen_clock(datetime(2024, 1, 1, 12, 0, 0))
    client = RecordingClient(
        pages=[[{"id": "mkt-1"}], [{"id": "mkt-2"}]],
        payloads={"mkt-1": {"name": "alpha"}, "mkt-2": {"name": "beta"}},
    )
    fetcher = MarketFetcher(client, default_limit=50, clock=clock)

    result = fetcher.fetch_markets(limit=500)

    assert set(result.keys()) == {"mkt-1", "mkt-2"}
    assert client.last_limit == 500
