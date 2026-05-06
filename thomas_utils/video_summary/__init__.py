"""Video-to-Markdown lecture summarizer.

Convert lecture videos into structured Markdown notes by combining
audio transcription (Whisper) with visual scene-change detection (OpenCV)
and multimodal LLM summarization (Claude / GPT-4o).
"""

from thomas_utils.video_summary.processor import VideoSummaryError, convert_video

__all__ = ["convert_video", "VideoSummaryError"]
