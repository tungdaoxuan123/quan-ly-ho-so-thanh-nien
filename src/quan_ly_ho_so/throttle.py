"""In-memory attempt limiter: blocks a key after too many hits inside a time window."""

import threading
import time
from collections import defaultdict, deque


class Throttle:
    def __init__(self, limit, window_seconds):
        self.limit = limit
        self.window = window_seconds
        self._hits = defaultdict(deque)
        self._lock = threading.Lock()

    def _prune(self, key, now):
        hits = self._hits[key]
        while hits and now - hits[0] > self.window:
            hits.popleft()
        return hits

    def blocked(self, key):
        with self._lock:
            return len(self._prune(key, time.monotonic())) >= self.limit

    def record(self, key):
        with self._lock:
            now = time.monotonic()
            self._prune(key, now).append(now)

    def clear(self, key):
        with self._lock:
            self._hits.pop(key, None)
