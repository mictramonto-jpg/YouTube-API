"""YouTube Data API v3 client for extracting video and comment data."""

import re
import time
import logging
from typing import Optional, Callable, List, Dict
from urllib.parse import urlparse, parse_qs
from datetime import datetime

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

logger = logging.getLogger(__name__)


class APIError(Exception):
    """YouTube API related errors with user-friendly messages."""

    _MESSAGES = {
        400: "リクエストが不正です。URLやパラメータを確認してください。",
        401: "API Keyが無効です。正しいキーを入力してください。",
        403: "APIのアクセスが拒否されました。APIキーの制限やクォータを確認してください。",
        404: "指定されたリソースが見つかりません。",
        429: "APIのリクエスト回数が上限に達しました。しばらく待ってから再度お試しください。",
    }

    def __init__(self, http_error: HttpError):
        self.status = http_error.resp.status
        detail = self._MESSAGES.get(self.status, f"APIエラーが発生しました (HTTP {self.status})")
        super().__init__(detail)
        self.original = http_error


def _api_call_with_retry(func, max_retries: int = 3, base_delay: float = 1.0):
    """Execute an API call with exponential backoff on transient errors."""
    for attempt in range(max_retries + 1):
        try:
            return func()
        except HttpError as e:
            status = e.resp.status
            # Non-retryable errors
            if status in (400, 401, 403, 404):
                raise
            # Retryable (429 quota, 500/503 server errors)
            if attempt < max_retries and status in (429, 500, 503):
                delay = base_delay * (2 ** attempt)
                logger.warning("API error %d, retry %d/%d in %.1fs",
                               status, attempt + 1, max_retries, delay)
                time.sleep(delay)
                continue
            raise


class YouTubeClient:
    """YouTube Data API v3 wrapper for video and comment extraction."""

    def __init__(self, api_key: str):
        self.api_key = api_key
        try:
            self.youtube = build("youtube", "v3", developerKey=api_key)
        except Exception as exc:
            raise ValueError(f"YouTube API の初期化に失敗しました: {exc}") from exc

    def validate_api_key(self) -> bool:
        """Make a cheap API call to verify the key works. Raises on failure."""
        try:
            _api_call_with_retry(
                lambda: self.youtube.videos().list(
                    part="id", id="dQw4w9WgXcQ", maxResults=1
                ).execute()
            )
            return True
        except HttpError as e:
            raise APIError(e) from e

    # ------------------------------------------------------------------
    # URL parsing
    # ------------------------------------------------------------------

    def parse_url(self, url: str) -> dict:
        """
        Parse a YouTube URL and return structured info.

        Returns dict with keys: type, id, tab
        """
        url = url.strip()
        if not url.startswith("http"):
            url = "https://" + url

        parsed = urlparse(url)
        path = parsed.path.rstrip("/")

        # Video URL: /watch?v=VIDEO_ID
        if "/watch" in path:
            params = parse_qs(parsed.query)
            video_id = params.get("v", [None])[0]
            if video_id:
                return {"type": "video", "id": video_id, "tab": "videos"}

        # Shorts video URL: /shorts/VIDEO_ID
        m = re.match(r"/shorts/([a-zA-Z0-9_-]+)", path)
        if m:
            return {"type": "video", "id": m.group(1), "tab": "shorts"}

        # Live video URL: /live/VIDEO_ID
        m = re.match(r"/live/([a-zA-Z0-9_-]+)", path)
        if m:
            return {"type": "video", "id": m.group(1), "tab": "streams"}

        # Handle URL: /@handle or /@handle/tab
        m = re.match(r"/@([^/]+)(?:/([^/]+))?", path)
        if m:
            return {
                "type": "handle",
                "id": m.group(1),
                "tab": _normalize_tab(m.group(2)),
            }

        # Channel ID URL: /channel/UCxxx or /channel/UCxxx/tab
        m = re.match(r"/channel/([^/]+)(?:/([^/]+))?", path)
        if m:
            return {
                "type": "channel_id",
                "id": m.group(1),
                "tab": _normalize_tab(m.group(2)),
            }

        # Custom URL: /c/name or /c/name/tab
        m = re.match(r"/c/([^/]+)(?:/([^/]+))?", path)
        if m:
            return {
                "type": "custom",
                "id": m.group(1),
                "tab": _normalize_tab(m.group(2)),
            }

        # User URL: /user/name
        m = re.match(r"/user/([^/]+)(?:/([^/]+))?", path)
        if m:
            return {
                "type": "username",
                "id": m.group(1),
                "tab": _normalize_tab(m.group(2)),
            }

        # Bare path: /name or /name/tab (legacy custom URLs)
        _reserved = {
            "watch", "shorts", "live", "channel", "c", "user",
            "playlist", "feed", "results", "gaming", "music",
            "premium", "account", "reporthistory", "hashtag",
        }
        m = re.match(r"/([a-zA-Z0-9_.-]+)(?:/([^/]+))?", path)
        if m and m.group(1).lower() not in _reserved:
            return {
                "type": "custom",
                "id": m.group(1),
                "tab": _normalize_tab(m.group(2)),
            }

        raise ValueError(
            "YouTubeのURLとして認識できません。\n"
            "対応形式: チャンネルURL (/@名前/videos) または 動画URL (/watch?v=...)"
        )

    # ------------------------------------------------------------------
    # Channel resolution
    # ------------------------------------------------------------------

    def resolve_channel_id(self, url_info: dict) -> str:
        """Resolve parsed URL info to a YouTube channel ID string."""
        info_type = url_info["type"]
        identifier = url_info["id"]

        try:
            if info_type == "channel_id":
                return identifier

            if info_type == "video":
                resp = _api_call_with_retry(
                    lambda: self.youtube.videos().list(
                        part="snippet", id=identifier
                    ).execute()
                )
                items = resp.get("items", [])
                if not items:
                    raise ValueError(f"動画が見つかりません: {identifier}")
                return items[0]["snippet"]["channelId"]

            if info_type == "handle":
                try:
                    resp = _api_call_with_retry(
                        lambda: self.youtube.channels().list(
                            part="id", forHandle=identifier
                        ).execute()
                    )
                    if resp.get("items"):
                        return resp["items"][0]["id"]
                except HttpError:
                    pass

            if info_type == "username":
                try:
                    resp = _api_call_with_retry(
                        lambda: self.youtube.channels().list(
                            part="id", forUsername=identifier
                        ).execute()
                    )
                    if resp.get("items"):
                        return resp["items"][0]["id"]
                except HttpError:
                    pass

            # Fallback: search for the channel
            if info_type in ("handle", "custom", "username"):
                query = f"@{identifier}" if info_type == "handle" else identifier
                resp = _api_call_with_retry(
                    lambda: self.youtube.search().list(
                        part="snippet", q=query, type="channel", maxResults=1
                    ).execute()
                )
                items = resp.get("items", [])
                if items:
                    return items[0]["snippet"]["channelId"]

        except HttpError as e:
            raise APIError(e) from e

        raise ValueError(f"チャンネルが見つかりません: {identifier}")

    # ------------------------------------------------------------------
    # Single video details (for single-video mode)
    # ------------------------------------------------------------------

    def get_video_details(self, video_id: str) -> Dict:
        """
        指定した1本の動画のメタデータ (タイトル・投稿日・投稿者名) を取得。

        Returns dict with: video_id, title, published_at, channel_title, url
        Raises ValueError if the video does not exist.
        """
        try:
            resp = _api_call_with_retry(
                lambda: self.youtube.videos().list(
                    part="snippet", id=video_id
                ).execute()
            )
        except HttpError as e:
            raise APIError(e) from e

        items = resp.get("items", [])
        if not items:
            raise ValueError(f"動画が見つかりません: {video_id}")

        snippet = items[0].get("snippet", {})
        return {
            "video_id": video_id,
            "title": snippet.get("title", ""),
            "published_at": snippet.get("publishedAt", ""),
            "channel_title": snippet.get("channelTitle", ""),
            "url": f"https://www.youtube.com/watch?v={video_id}",
        }

    # ------------------------------------------------------------------
    # Video listing
    # ------------------------------------------------------------------

    def get_videos(
        self,
        channel_id: str,
        tab: str,
        max_videos: int,
        order: str = "date",
        progress_callback: Optional[Callable[[str], None]] = None,
        cancel_check: Optional[Callable[[], bool]] = None,
    ) -> List[Dict]:
        """
        Get videos from a channel.

        Args:
            order: 並び順
                - "date"      : 新しい順 (デフォルト)
                - "viewCount" : 視聴回数順 (再生が多い順)
                - "rating"    : 評価順
                - "oldest"    : 古い順
                - "relevance" : 関連度順 (YouTubeアルゴリズム推奨順)
        """
        # "videos" タブで「新しい順」の場合のみ uploads playlist が最安
        # (playlistItems.list = 1 unit)。その他はすべて search.list (100 units/page)
        if tab == "videos" and order == "date":
            return self._get_uploads(
                channel_id, max_videos,
                progress_callback=progress_callback,
                cancel_check=cancel_check,
            )

        # 以降は search.list を使用
        search_order = order if order != "oldest" else "date"
        reverse = (order == "oldest")

        video_duration = "short" if tab == "shorts" else None
        event_type = "completed" if tab == "streams" else None

        results = self._search_videos(
            channel_id, max_videos,
            video_duration=video_duration,
            event_type=event_type,
            order=search_order,
            progress_callback=progress_callback,
            cancel_check=cancel_check,
        )

        # 「古い順」は date の結果を末尾から取得し直す必要がある
        # ただし search.list は最大500件までしか取得できないため、
        # 完全な古い順にはならない場合がある
        if reverse:
            results.reverse()

        return results

    def _get_uploads(self, channel_id, max_videos,
                     progress_callback=None, cancel_check=None):
        """Get videos from the channel's uploads playlist (newest first)."""
        if progress_callback:
            progress_callback("チャンネル情報を取得中...")

        try:
            ch_resp = _api_call_with_retry(
                lambda: self.youtube.channels().list(
                    part="contentDetails", id=channel_id
                ).execute()
            )
        except HttpError as e:
            raise APIError(e) from e

        items = ch_resp.get("items", [])
        if not items:
            raise ValueError(f"チャンネルが見つかりません: {channel_id}")

        uploads_id = items[0]["contentDetails"]["relatedPlaylists"]["uploads"]

        videos = []  # type: List[Dict]
        next_page = None
        _skip_titles = {
            "Deleted video", "Private video", "削除された動画", "非公開動画",
        }

        while len(videos) < max_videos:
            if cancel_check and cancel_check():
                break

            if progress_callback:
                progress_callback(f"動画リスト取得中... ({len(videos)}/{max_videos})")

            try:
                _next = next_page  # capture for lambda
                resp = _api_call_with_retry(
                    lambda: self.youtube.playlistItems().list(
                        part="snippet",
                        playlistId=uploads_id,
                        maxResults=min(50, max_videos - len(videos)),
                        pageToken=_next,
                    ).execute()
                )
            except HttpError as e:
                raise APIError(e) from e

            for item in resp.get("items", []):
                snippet = item.get("snippet", {})
                title = snippet.get("title", "")
                if title in _skip_titles:
                    continue
                resource = snippet.get("resourceId", {})
                video_id = resource.get("videoId", "")
                if not video_id:
                    continue
                videos.append({
                    "video_id": video_id,
                    "title": title,
                    "published_at": snippet.get("publishedAt", ""),
                    "channel_title": snippet.get("channelTitle", ""),
                    "url": f"https://www.youtube.com/watch?v={video_id}",
                })
                if len(videos) >= max_videos:
                    break

            next_page = resp.get("nextPageToken")
            if not next_page:
                break

        return videos[:max_videos]

    def _search_videos(self, channel_id, max_videos,
                       video_duration=None, event_type=None,
                       order="date",
                       progress_callback=None, cancel_check=None):
        """Get videos via the Search API (for shorts / streams)."""
        videos = []  # type: List[Dict]
        next_page = None

        while len(videos) < max_videos:
            if cancel_check and cancel_check():
                break

            if progress_callback:
                progress_callback(f"動画リスト取得中... ({len(videos)}/{max_videos})")

            params = {
                "part": "snippet",
                "channelId": channel_id,
                "type": "video",
                "order": order,
                "maxResults": min(50, max_videos - len(videos)),
            }
            if video_duration:
                params["videoDuration"] = video_duration
            if event_type:
                params["eventType"] = event_type
            if next_page:
                params["pageToken"] = next_page

            try:
                resp = _api_call_with_retry(
                    lambda: self.youtube.search().list(**params).execute()
                )
            except HttpError as e:
                raise APIError(e) from e

            for item in resp.get("items", []):
                vid = item.get("id", {})
                video_id = vid.get("videoId", "")
                if not video_id:
                    continue
                snippet = item.get("snippet", {})
                videos.append({
                    "video_id": video_id,
                    "title": snippet.get("title", ""),
                    "published_at": snippet.get("publishedAt", ""),
                    "channel_title": snippet.get("channelTitle", ""),
                    "url": f"https://www.youtube.com/watch?v={video_id}",
                })
                if len(videos) >= max_videos:
                    break

            next_page = resp.get("nextPageToken")
            if not next_page:
                break

        return videos[:max_videos]

    # ------------------------------------------------------------------
    # Comment extraction
    # ------------------------------------------------------------------

    def get_video_comments(
        self,
        video_id: str,
        min_likes: int = 0,
        text_filter: str = "",
        cancel_check: Optional[Callable[[], bool]] = None,
    ) -> List[Dict]:
        """
        Retrieve top-level comments for a video, ordered by relevance.

        Filters are applied on-the-fly:
          - min_likes: only include comments with >= this many likes
          - text_filter: only include comments whose text contains this string
        """
        comments = []  # type: List[Dict]
        next_page = None

        while True:
            if cancel_check and cancel_check():
                break

            try:
                _next = next_page
                resp = _api_call_with_retry(
                    lambda: self.youtube.commentThreads().list(
                        part="snippet",
                        videoId=video_id,
                        order="relevance",
                        maxResults=100,
                        pageToken=_next,
                    ).execute()
                )
            except HttpError as e:
                status = e.resp.status
                if status in (403, 404):
                    # 403 = comments disabled, 404 = video not found
                    logger.info("Skipping comments for %s (HTTP %d)", video_id, status)
                    break
                raise APIError(e) from e

            for item in resp.get("items", []):
                try:
                    snippet = item["snippet"]["topLevelComment"]["snippet"]
                except (KeyError, TypeError):
                    continue

                like_count = snippet.get("likeCount", 0)
                text = snippet.get("textOriginal", "")

                if min_likes > 0 and like_count < min_likes:
                    continue
                if text_filter and text_filter not in text:
                    continue

                comments.append({
                    "author": snippet.get("authorDisplayName", ""),
                    "text": text,
                    "likes": like_count,
                    "published_at": snippet.get("publishedAt", ""),
                })

            next_page = resp.get("nextPageToken")
            if not next_page:
                break

        return comments


# ------------------------------------------------------------------
# Date / time formatting helpers
# ------------------------------------------------------------------

def format_date(iso_string: str) -> str:
    """ISO 8601 -> 'YYYY/MM/DD'."""
    if not iso_string:
        return ""
    try:
        dt = datetime.fromisoformat(iso_string.replace("Z", "+00:00"))
        return dt.strftime("%Y/%m/%d")
    except (ValueError, AttributeError):
        return iso_string


def format_datetime(iso_string: str) -> str:
    """ISO 8601 -> 'YYYY/MM/DD HH:MM:SS'."""
    if not iso_string:
        return ""
    try:
        dt = datetime.fromisoformat(iso_string.replace("Z", "+00:00"))
        return dt.strftime("%Y/%m/%d %H:%M:%S")
    except (ValueError, AttributeError):
        return iso_string


def estimate_quota_level(
    num_urls: int,
    mode: str = "video",
    max_videos_per_channel: int = 30,
    order: str = "date",
    tab_hint: str = "videos",
) -> str:
    """
    入力条件から想定APIクォータ消費量を推定し、定性的なレベルを返す。

    注: YouTube Data APIは実際の残量を取得できないため、これは目安です。
    コメント数は動画により大きく異なるため、実消費は想定と乖離します。

    Returns: "少" | "中" | "多" | "大"
    """
    # 内部で大まかなユニット数を計算 (表示はしない)
    cost_est = 1  # APIキー検証

    if mode == "video":
        # videos.list × N + commentThreads.list × N × avg_pages
        cost_est += num_urls * 1
        cost_est += num_urls * 5  # 平均5ページ想定
    else:
        # チャンネル特定 (× N)
        cost_est += num_urls * 1
        # 動画リスト取得
        pages = max(1, (max_videos_per_channel + 49) // 50)
        if tab_hint == "videos" and order == "date":
            cost_est += num_urls * pages * 1
        else:
            cost_est += num_urls * pages * 100  # search.list
        # コメント取得
        total_videos = num_urls * max_videos_per_channel
        cost_est += total_videos * 5

    # 無料枠 10,000 units に対する割合でレベル判定
    # 少: ~5% (< 500)   中: 5~20% (500~2000)
    # 多: 20~50% (2000~5000)   大: >50% (5000+)
    if cost_est < 500:
        return "少"
    if cost_est < 2000:
        return "中"
    if cost_est < 5000:
        return "多"
    return "大"


def _normalize_tab(tab: Optional[str]) -> str:
    """Normalize a URL path segment to a canonical tab name."""
    if not tab:
        return "videos"
    tab = tab.lower().strip("/")
    if tab in ("videos", "video", "featured"):
        return "videos"
    if tab in ("shorts", "short"):
        return "shorts"
    if tab in ("streams", "stream", "live"):
        return "streams"
    return "videos"
