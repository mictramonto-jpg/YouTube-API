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

    # YouTube Data API v3 のエンドポイント別クォータコスト (公式値)
    # https://developers.google.com/youtube/v3/determine_quota_cost
    QUOTA_COSTS = {
        "videos.list":         1,
        "channels.list":       1,
        "playlistItems.list":  1,
        "commentThreads.list": 1,
        "search.list":       100,
    }

    # デフォルトの1日あたり無料クォータ (Google Cloud Console で申請すれば増量可)
    DAILY_FREE_QUOTA = 10_000

    def __init__(self, api_key: str, quota_callback: Optional[Callable[[int, str], None]] = None):
        """
        Args:
            api_key: YouTube Data API v3 Key
            quota_callback: API呼び出しのたびに (累計消費ユニット, 最後の呼び出し名) を
                            通知するコールバック (GUI側の表示更新に使用)
        """
        self.api_key = api_key
        self.quota_used = 0  # セッション内累計消費ユニット
        self.quota_callback = quota_callback
        try:
            self.youtube = build("youtube", "v3", developerKey=api_key)
        except Exception as exc:
            raise ValueError(f"YouTube API の初期化に失敗しました: {exc}") from exc

    def _track_quota(self, endpoint: str):
        """API呼び出し時にクォータ消費を記録。"""
        cost = self.QUOTA_COSTS.get(endpoint, 1)
        self.quota_used += cost
        if self.quota_callback:
            try:
                self.quota_callback(self.quota_used, endpoint)
            except Exception:
                pass  # コールバックの失敗はAPI処理に影響させない

    def _call(self, endpoint: str, func):
        """クォータ追跡付きのAPI呼び出しラッパー。"""
        result = _api_call_with_retry(func)
        self._track_quota(endpoint)
        return result

    def validate_api_key(self) -> bool:
        """Make a cheap API call to verify the key works. Raises on failure."""
        try:
            self._call(
                "videos.list",
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
                resp = self._call(
                    "videos.list",
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
                    resp = self._call(
                        "channels.list",
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
                    resp = self._call(
                        "channels.list",
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
                resp = self._call(
                    "search.list",
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
            resp = self._call(
                "videos.list",
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
            ch_resp = self._call(
                "channels.list",
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
                resp = self._call(
                    "playlistItems.list",
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
                resp = self._call(
                    "search.list",
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
                resp = self._call(
                    "commentThreads.list",
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
# Quota estimation (pre-run cost calculator)
# ------------------------------------------------------------------

def estimate_quota_cost(
    num_urls: int,
    mode: str = "video",
    max_videos_per_channel: int = 30,
    order: str = "date",
    tab_hint: str = "videos",
    avg_comments_per_video: int = 500,
) -> Dict:
    """
    入力条件から、1回の稼働でおおよそ消費するクォータ量を見積もる。

    Returns:
        {
            "total": 合計ユニット,
            "breakdown": [(ラベル, ユニット), ...] 内訳,
            "per_video_avg": 動画1本あたりの平均消費,
        }
    """
    breakdown = []
    total = 0

    # 1. APIキー検証 (videos.list × 1)
    breakdown.append(("APIキー検証 (videos.list)", 1))
    total += 1

    if mode == "video":
        # 各URLについて動画詳細を1回取得
        cost = num_urls * 1
        breakdown.append((f"動画情報取得 (videos.list × {num_urls})", cost))
        total += cost

        # コメント取得: 1動画あたり平均 ceil(avg/100) ページ
        pages_per_video = max(1, (avg_comments_per_video + 99) // 100)
        cost = num_urls * pages_per_video * 1
        breakdown.append((
            f"コメント取得 (commentThreads.list × {num_urls}動画 × {pages_per_video}ページ)",
            cost,
        ))
        total += cost

    else:  # channel mode
        # 各URLごとにチャンネル特定
        cost = num_urls * 1  # channels.list or videos.list
        breakdown.append((f"チャンネル特定 (channels.list × {num_urls})", cost))
        total += cost

        # 動画リスト取得
        pages = max(1, (max_videos_per_channel + 49) // 50)
        if tab_hint == "videos" and order == "date":
            # uploads playlist (1 unit/page)
            cost = num_urls * pages * 1
            breakdown.append((
                f"動画リスト取得 (playlistItems.list × {num_urls}ch × {pages}p)",
                cost,
            ))
        else:
            # search.list (100 units/page) - 高コスト
            cost = num_urls * pages * 100
            breakdown.append((
                f"動画リスト取得 ⚠️ search.list × {num_urls}ch × {pages}p = 100units/req",
                cost,
            ))
        total += cost

        # コメント取得
        total_videos = num_urls * max_videos_per_channel
        pages_per_video = max(1, (avg_comments_per_video + 99) // 100)
        cost = total_videos * pages_per_video * 1
        breakdown.append((
            f"コメント取得 (commentThreads.list × {total_videos}動画 × {pages_per_video}p)",
            cost,
        ))
        total += cost

    if mode == "video":
        per_video = total / max(1, num_urls)
    else:
        per_video = total / max(1, num_urls * max_videos_per_channel)

    return {
        "total": total,
        "breakdown": breakdown,
        "per_video_avg": per_video,
    }


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
