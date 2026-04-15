import streamlit as st
import yt_dlp
import os
import re
import time
import tempfile
import io
import base64
import socket
import asyncio
import cv2
import requests
import numpy as np
from PIL import Image, ImageEnhance, ImageFilter
from deep_translator import GoogleTranslator
import edge_tts

# ── Page config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="K-POP 韓語學習簡報產生器",
    page_icon="🎤",
    layout="centered",
)

# ── Custom CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-title {
        font-size: 2.2rem; font-weight: 800; text-align: center;
        background: linear-gradient(135deg, #FF6B9D, #C44B8A, #9B59B6);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        margin-bottom: 0.2rem;
    }
    .sub-title { text-align: center; color: #888; font-size: 0.95rem; margin-bottom: 2rem; }
    .stButton > button {
        width: 100%;
        background: linear-gradient(135deg, #FF6B9D, #9B59B6);
        color: white; border: none; border-radius: 8px;
        padding: 0.6rem 1.2rem; font-size: 1rem; font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">K-POP 韓語學習簡報產生器 🎤</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">貼上 YouTube MV 網址，自動產生韓中對照投影片（含語音）</div>', unsafe_allow_html=True)

# ── Input ────────────────────────────────────────────────────────────────────
url = st.text_input("YouTube 網址", placeholder="https://www.youtube.com/watch?v=...",
                    label_visibility="collapsed")

def _is_cloud() -> bool:
    try:
        h = socket.gethostname()
        return "streamlit" in h.lower() or h.startswith("ip-")
    except Exception:
        return False

_running_on_cloud = _is_cloud()

col1, col2 = st.columns(2)
with col1:
    if _running_on_cloud:
        use_mv_frames = False
        st.info("☁️ 雲端版使用縮圖背景")
    else:
        use_mv_frames = st.checkbox("🎬 擷取 MV 畫面", value=True,
                                    help="下載影片擷取對應畫面（較慢）")
with col2:
    use_tts = st.checkbox("🔊 加入韓文語音", value=True,
                          help="每頁自動朗讀韓文歌詞")

# ═══════════════════════════════════════════════════════════════════════════════
# Helper functions
# ═══════════════════════════════════════════════════════════════════════════════

def get_video_id(url: str) -> str | None:
    m = re.search(r"(?:v=|youtu\.be/)([a-zA-Z0-9_-]{11})", url)
    return m.group(1) if m else None

def get_video_info(url: str) -> dict:
    ydl_opts = {"quiet": True, "no_warnings": True, "skip_download": True, "socket_timeout": 15}
    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(url, download=False)
            if info is None:
                raise ValueError("影片不存在或無法存取")
            return info
    except SystemExit as e:
        raise ValueError(f"影片可能不存在或有地區限制：{e}")

def download_subtitles(url: str, tmpdir: str) -> tuple[str | None, str]:
    base = os.path.join(tmpdir, "subtitle")
    for auto in (False, True):
        ydl_opts = {
            "quiet": True, "no_warnings": True, "skip_download": True,
            "writesubtitles": not auto, "writeautomaticsub": auto,
            "subtitleslangs": ["ko"], "subtitlesformat": "vtt", "outtmpl": base,
        }
        try:
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                ydl.download([url])
        except (SystemExit, Exception):
            pass
        for ext in ("vtt", "srt"):
            candidate = f"{base}.ko.{ext}"
            if os.path.exists(candidate):
                return candidate, "自動生成" if auto else "官方提供"
    return None, "找不到韓文字幕"

def _ts_to_sec(ts: str) -> float:
    parts = ts.strip().split(":")
    if len(parts) == 3:
        h, m, s = int(parts[0]), int(parts[1]), float(parts[2])
    else:
        h, m, s = 0, int(parts[0]), float(parts[1])
    return h * 3600 + m * 60 + s

def parse_vtt(path: str) -> list[tuple[float, str]]:
    with open(path, encoding="utf-8") as f:
        content = f.read()
    results, last_text = [], ""
    for block in re.split(r"\n\s*\n", content.strip()):
        lines = block.strip().splitlines()
        ts_sec, text_parts = None, []
        for line in lines:
            ts_m = re.match(r"([\d:]+\.[\d]+)\s*-->", line)
            if ts_m:
                ts_sec = _ts_to_sec(ts_m.group(1)); continue
            if re.match(r"^\d+$", line.strip()) or re.match(r"WEBVTT|NOTE", line): continue
            clean = re.sub(r"<[^>]+>", "", line).strip()
            if clean: text_parts.append(clean)
        if ts_sec is not None and text_parts:
            text = " ".join(text_parts)
            if text != last_text:
                last_text = text
                results.append((ts_sec, text))
    return results

def parse_srt(path: str) -> list[tuple[float, str]]:
    with open(path, encoding="utf-8") as f:
        content = f.read()
    results, last_text = [], ""
    for block in re.split(r"\n\s*\n", content.strip()):
        lines = block.strip().splitlines()
        ts_sec, text_parts = None, []
        for line in lines:
            if re.match(r"^\d+$", line.strip()): continue
            ts_m = re.match(r"([\d:,]+)\s*-->", line)
            if ts_m:
                ts_sec = _ts_to_sec(ts_m.group(1).replace(",", ".")); continue
            clean = re.sub(r"<[^>]+>", "", line).strip()
            if clean: text_parts.append(clean)
        if ts_sec is not None and text_parts:
            text = " ".join(text_parts)
            if text != last_text:
                last_text = text
                results.append((ts_sec, text))
    return results

def translate_with_retry(text: str, retries: int = 4, delay: float = 1.5) -> str:
    for attempt in range(retries):
        try:
            result = GoogleTranslator(source="ko", target="zh-TW").translate(text)
            if result: return result
        except Exception:
            pass
        time.sleep(delay * (attempt + 1))
    return ""

def translate_lines(entries: list[tuple[float, str]], status_ph) -> list[str]:
    translated, total = [], len(entries)
    for i, (_, ko) in enumerate(entries, 1):
        status_ph.text(f"翻譯中... ({i}/{total})")
        translated.append(translate_with_retry(ko))
        time.sleep(0.1)
    return translated

def _ffmpeg_available() -> bool:
    import shutil
    return shutil.which("ffmpeg") is not None

def download_video(url: str, tmpdir: str, status_ph) -> str | None:
    out_tmpl = os.path.join(tmpdir, "video.%(ext)s")
    fmt = "bestvideo[height<=720][ext=mp4]/bestvideo[height<=720]/bestvideo[height<=480][ext=mp4]/bestvideo[height<=480]/worst[ext=mp4]/worst"
    ydl_opts = {"quiet": True, "no_warnings": True, "format": fmt, "outtmpl": out_tmpl, "noplaylist": True}
    status_ph.text("正在下載 MV 影片…")
    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            ydl.extract_info(url, download=True)
        for f in os.listdir(tmpdir):
            if f.startswith("video."):
                return os.path.join(tmpdir, f)
    except Exception as e:
        status_ph.text(f"影片下載失敗：{e}")
    return None

def get_thumbnail(video_id: str) -> bytes | None:
    for quality in ("maxresdefault", "hqdefault", "mqdefault"):
        try:
            resp = requests.get(f"https://img.youtube.com/vi/{video_id}/{quality}.jpg", timeout=10)
            if resp.status_code == 200 and len(resp.content) > 5000:
                return resp.content
        except Exception:
            pass
    return None

def _process_frame_image(img: Image.Image, brightness: float = 0.60) -> bytes:
    img = img.convert("RGB").resize((1280, 720), Image.LANCZOS)
    img = ImageEnhance.Brightness(img).enhance(brightness)
    img = img.filter(ImageFilter.GaussianBlur(radius=0.8))
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=92)
    buf.seek(0)
    return buf.read()

def extract_all_frames(video_path: str, timestamps: list[float],
                       brightness: float = 0.60, status_ph=None) -> dict[float, bytes]:
    frames: dict[float, bytes] = {}
    if not video_path or not os.path.exists(video_path): return frames
    try:
        cap = cv2.VideoCapture(video_path)
        total = len(timestamps)
        for i, ts in enumerate(timestamps):
            if status_ph: status_ph.text(f"擷取 MV 畫面... ({i+1}/{total})")
            cap.set(cv2.CAP_PROP_POS_MSEC, ts * 1000)
            ret, frame_bgr = cap.read()
            if ret:
                frame_rgb = cv2.cvtColor(frame_bgr, cv2.COLOR_BGR2RGB)
                frames[ts] = _process_frame_image(Image.fromarray(frame_rgb), brightness)
        cap.release()
    except Exception:
        pass
    return frames

def prepare_bg_image(img_bytes: bytes | None, brightness: float = 0.60) -> bytes | None:
    if not img_bytes: return None
    try:
        img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
        img = img.resize((1280, 720), Image.LANCZOS)
        img = ImageEnhance.Brightness(img).enhance(brightness)
        img = img.filter(ImageFilter.GaussianBlur(radius=0.8))
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=88)
        buf.seek(0)
        return buf.read()
    except Exception:
        return None

async def _tts_async(text: str, voice: str) -> bytes:
    communicate = edge_tts.Communicate(text, voice)
    buf = io.BytesIO()
    async for chunk in communicate.stream():
        if chunk["type"] == "audio":
            buf.write(chunk["data"])
    buf.seek(0)
    return buf.read()

def generate_tts(text: str, voice: str = "ko-KR-SunHiNeural") -> bytes | None:
    """Microsoft Edge TTS — ko-KR-SunHiNeural is a natural female Korean voice."""
    try:
        loop = asyncio.new_event_loop()
        try:
            return loop.run_until_complete(_tts_async(text, voice))
        finally:
            loop.close()
    except Exception:
        return None

def b64(data: bytes, mime: str) -> str:
    return f"data:{mime};base64,{base64.b64encode(data).decode()}"

# ═══════════════════════════════════════════════════════════════════════════════
# HTML Builder
# ═══════════════════════════════════════════════════════════════════════════════

def build_html(title: str,
               entries: list[tuple[float, str]],
               chinese_lines: list[str],
               thumb_bytes: bytes | None,
               frames: dict[float, bytes],
               tts_map: dict[int, bytes],
               status_ph) -> str:

    thumb_bg = prepare_bg_image(thumb_bytes, brightness=0.50)
    cover_bg_css = f"background-image:url('{b64(thumb_bg, 'image/jpeg')}');background-size:cover;background-position:center;" if thumb_bg else "background:#1A0533;"

    slides_js = []
    total = len(entries)

    # Cover slide
    slides_js.append(f"""{{
  type:'cover',
  title:{repr(title)},
  bg:{repr(b64(thumb_bg,'image/jpeg') if thumb_bg else '')},
  audio:''
}}""")

    for idx, ((ts, ko), zh) in enumerate(zip(entries, chinese_lines)):
        status_ph.text(f"生成 HTML... ({idx+1}/{total})")

        # Background
        frame = frames.get(ts) or thumb_bg
        bg_b64 = b64(frame, 'image/jpeg') if frame else ''

        # Audio
        audio_b64 = ''
        if idx in tts_map and tts_map[idx]:
            audio_b64 = b64(tts_map[idx], 'audio/mp3')

        slides_js.append(f"""{{
  type:'lyric',
  num:{idx+1},
  ko:{repr(ko)},
  zh:{repr(zh if zh else ko)},
  bg:{repr(bg_b64)},
  audio:{repr(audio_b64)}
}}""")

    slides_json = ",\n".join(slides_js)

    html = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>{title}</title>
<style>
  * {{ margin:0; padding:0; box-sizing:border-box; }}
  body {{ background:#000; font-family:'Noto Sans KR','Apple SD Gothic Neo','Microsoft JhengHei',sans-serif; overflow:hidden; height:100vh; }}

  #deck {{ width:100%; height:100vh; position:relative; }}

  .slide {{
    position:absolute; inset:0;
    display:flex; flex-direction:column; justify-content:center; align-items:center;
    opacity:0; pointer-events:none; transition:opacity 0.5s ease;
    background:#1A0533;
  }}
  .slide.active {{ opacity:1; pointer-events:all; }}
  .slide-bg {{
    position:absolute; inset:0;
    background-size:cover; background-position:center;
    z-index:0;
  }}
  .slide-content {{ position:relative; z-index:1; text-align:center; padding:2rem; width:100%; }}

  /* Cover */
  .cover-badge {{ color:#C490D1; font-size:1rem; letter-spacing:0.2em; margin-bottom:1.5rem; font-style:italic; }}
  .cover-title {{ color:#fff; font-size:clamp(1.8rem,5vw,3.2rem); font-weight:800; line-height:1.3; margin-bottom:1.5rem; text-shadow:0 2px 20px rgba(0,0,0,0.8); }}
  .cover-dots {{ color:#FF6B9D; font-size:1.5rem; margin-bottom:1rem; }}
  .cover-sub {{ color:#C490D1; font-size:0.9rem; }}

  /* Lyric */
  .slide-num {{ position:absolute; top:1rem; left:1.2rem; color:#FF6B9D; font-size:0.9rem; font-weight:700; }}
  .ko-text {{ color:#fff; font-size:clamp(1.4rem,4.5vw,2.8rem); font-weight:700; line-height:1.5; margin-bottom:1.5rem; text-shadow:0 2px 20px rgba(0,0,0,0.9); }}
  .divider {{ width:60px; height:2px; background:#FF6B9D; margin:0 auto 1.5rem; }}
  .zh-text {{ color:#F9C6DD; font-size:clamp(1rem,3vw,1.6rem); line-height:1.6; text-shadow:0 2px 15px rgba(0,0,0,0.9); }}

  /* Top/bottom bars */
  .bar-top,.bar-bottom {{ position:absolute; left:0; right:0; height:5px; background:#FF6B9D; z-index:2; }}
  .bar-top {{ top:0; }} .bar-bottom {{ bottom:0; }}

  /* Audio buttons — centered above nav */
  .audio-btns {{
    position:absolute; bottom:5.2rem; left:50%; transform:translateX(-50%);
    display:flex; gap:0.7rem; z-index:3; white-space:nowrap;
  }}
  .audio-btn {{
    background:rgba(255,107,157,0.82); border:none; border-radius:24px;
    padding:0.45rem 1.1rem; cursor:pointer; font-size:1rem; color:#fff;
    display:flex; align-items:center; gap:0.4rem; font-weight:600;
    transition:transform 0.2s, background 0.2s; box-shadow:0 2px 10px rgba(0,0,0,0.4);
    backdrop-filter:blur(6px);
  }}
  .audio-btn:hover {{ transform:scale(1.06); background:rgba(255,107,157,1); }}
  .audio-btn:disabled {{ opacity:0.5; cursor:default; transform:none; }}

  /* Nav */
  #nav {{
    position:fixed; bottom:1.2rem; left:50%; transform:translateX(-50%);
    display:flex; gap:1rem; align-items:center; z-index:10;
  }}
  #nav button {{
    background:rgba(255,255,255,0.15); border:1px solid rgba(255,255,255,0.3);
    color:#fff; padding:0.5rem 1.2rem; border-radius:20px; cursor:pointer;
    font-size:0.9rem; backdrop-filter:blur(8px); transition:background 0.2s;
  }}
  #nav button:hover {{ background:rgba(255,107,157,0.5); }}
  #counter {{ color:rgba(255,255,255,0.6); font-size:0.85rem; min-width:60px; text-align:center; }}

  /* Progress bar */
  #progress {{ position:fixed; top:0; left:0; height:3px; background:#FF6B9D; z-index:20; transition:width 0.3s; }}

  /* Speed control */
  #speed-ctrl {{
    position:fixed; bottom:4.2rem; left:50%; transform:translateX(-50%);
    display:flex; gap:0.4rem; align-items:center; z-index:10;
  }}
  #speed-ctrl button {{
    background:rgba(255,255,255,0.10); border:1px solid rgba(255,255,255,0.22);
    color:#bbb; padding:0.28rem 0.75rem; border-radius:14px; cursor:pointer;
    font-size:0.78rem; backdrop-filter:blur(8px); transition:all 0.18s;
  }}
  #speed-ctrl button.sp-active {{
    background:rgba(255,107,157,0.55); border-color:#FF6B9D; color:#fff; font-weight:700;
  }}
  #speed-ctrl button:hover {{ background:rgba(255,107,157,0.30); color:#fff; }}
</style>
</head>
<body>
<div id="progress"></div>
<div id="deck"></div>
<div id="speed-ctrl">
  <button onclick="setSpeed(0.9,this)">0.9×</button>
  <button onclick="setSpeed(1.0,this)" class="sp-active">1×</button>
  <button onclick="setSpeed(1.1,this)">1.1×</button>
</div>
<div id="nav">
  <button onclick="go(-1)">&#8592;</button>
  <span id="counter">1 / 1</span>
  <button onclick="go(1)">&#8594;</button>
</div>

<script>
const slides = [
{slides_json}
];

let cur = 0;
let audioEl = null;
let playbackRate = 1.0;

function setSpeed(rate, el){{
  playbackRate = rate;
  document.querySelectorAll('#speed-ctrl button').forEach(b => b.classList.remove('sp-active'));
  el.classList.add('sp-active');
  if(audioEl && !audioEl.paused){{
    // Already playing — just change the rate live
    audioEl.playbackRate = rate;
  }} else {{
    // Not playing — replay current slide so user can hear the difference immediately
    replayCurrentSlide();
  }}
}}

function _makeAudio(b64){{
  const a = new Audio(b64);
  a.playbackRate = playbackRate;
  a.addEventListener('canplay', ()=>{{ a.playbackRate = playbackRate; }}, {{once:true}});
  return a;
}}

function playAudio(b64){{
  if(!b64) return;
  if(audioEl){{ audioEl.pause(); audioEl=null; }}
  audioEl = _makeAudio(b64);
  audioEl.play().catch(()=>{{}});
}}

let _repeat3Timer = null;
function playAudio3x(b64, btns){{
  if(!b64) return;
  // Disable buttons while playing
  if(btns) btns.forEach(b=>b.disabled=true);
  if(_repeat3Timer) clearTimeout(_repeat3Timer);
  if(audioEl){{ audioEl.pause(); audioEl=null; }}
  let count = 0;
  function playOne(){{
    if(audioEl){{ audioEl.pause(); audioEl=null; }}
    const a = _makeAudio(b64);
    audioEl = a;
    a.addEventListener('ended', ()=>{{
      audioEl = null;
      count++;
      if(count < 3){{
        _repeat3Timer = setTimeout(playOne, 500);
      }} else {{
        if(btns) btns.forEach(b=>b.disabled=false);
      }}
    }});
    a.play().catch(()=>{{ if(btns) btns.forEach(b=>b.disabled=false); }});
  }}
  playOne();
}}

function replayCurrentSlide(){{
  if(slides[cur] && slides[cur].audio) playAudio(slides[cur].audio);
}}

function renderSlides(){{
  const deck = document.getElementById('deck');
  deck.innerHTML = '';
  slides.forEach((s,i)=>{{
    const el = document.createElement('div');
    el.className = 'slide' + (i===0?' active':'');
    el.id = 'slide'+i;

    // Background
    if(s.bg){{
      const bg = document.createElement('div');
      bg.className = 'slide-bg';
      bg.style.backgroundImage = `url('${{s.bg}}')`;
      el.appendChild(bg);
    }}

    const bars = '<div class="bar-top"></div><div class="bar-bottom"></div>';

    if(s.type==='cover'){{
      el.innerHTML += bars + `
        <div class="slide-content">
          <div class="cover-badge">K-POP 韓語學習</div>
          <div class="cover-title">${{s.title}}</div>
          <div class="cover-dots">· · ·</div>
          <div class="cover-sub">韓中對照學習投影片 / 한중 대조 학습 슬라이드</div>
        </div>`;
    }} else {{
      const audioBtns = s.audio ? `
        <div class="audio-btns" id="abtns${{i}}">
          <button class="audio-btn" onclick="playAudio(slides[${{i}}].audio)" title="播放一次">▶ 播放</button>
          <button class="audio-btn" onclick="playAudio3x(slides[${{i}}].audio,[...document.querySelectorAll('#abtns${{i}} button')])" title="播放三次">🔁 ×3</button>
        </div>` : '';
      el.innerHTML += bars + `
        <div class="slide-num">${{String(s.num).padStart(2,'0')}}</div>
        <div class="slide-content">
          <div class="ko-text">${{s.ko}}</div>
          <div class="divider"></div>
          <div class="zh-text">${{s.zh}}</div>
        </div>
        ${{audioBtns}}`;
    }}
    deck.appendChild(el);
  }});
}}

function go(dir){{
  const prev = cur;
  cur = Math.max(0, Math.min(slides.length-1, cur+dir));
  if(cur===prev) return;
  document.getElementById('slide'+prev).classList.remove('active');
  document.getElementById('slide'+cur).classList.add('active');
  document.getElementById('counter').textContent = (cur+1)+' / '+slides.length;
  document.getElementById('progress').style.width = ((cur/(slides.length-1))*100)+'%';
  // Auto-play audio on slide change
  if(slides[cur].audio) playAudio(slides[cur].audio);
  else if(audioEl){{ audioEl.pause(); audioEl=null; }}
}}

// Keyboard navigation
document.addEventListener('keydown', e=>{{
  if(e.key==='ArrowRight'||e.key==='ArrowDown') go(1);
  if(e.key==='ArrowLeft'||e.key==='ArrowUp') go(-1);
}});

// Touch swipe
let tx=0;
document.addEventListener('touchstart',e=>{{ tx=e.touches[0].clientX; }});
document.addEventListener('touchend',e=>{{
  const dx=tx-e.changedTouches[0].clientX;
  if(Math.abs(dx)>50) go(dx>0?1:-1);
}});

renderSlides();
document.getElementById('counter').textContent = '1 / '+slides.length;
document.getElementById('progress').style.width = '0%';
</script>
</body>
</html>"""

    return html

# ═══════════════════════════════════════════════════════════════════════════════
# Main flow
# ═══════════════════════════════════════════════════════════════════════════════

if st.button("🎵 開始轉換"):
    if not url.strip():
        st.error("請輸入 YouTube 網址。"); st.stop()
    if not re.match(r"https?://(www\.)?(youtube\.com|youtu\.be)/", url.strip()):
        st.error("請輸入有效的 YouTube 網址。"); st.stop()

    vid = get_video_id(url.strip())
    if vid:
        url = f"https://www.youtube.com/watch?v={vid}"

    html_bytes: bytes | None = None
    video_title = "K-POP 影片"

    with tempfile.TemporaryDirectory() as tmpdir:

        # Step 1: Video info
        with st.spinner("正在取得影片資訊…"):
            try:
                info = get_video_info(url)
                video_title = info.get("title", "K-POP 影片")
                video_id    = get_video_id(url) or info.get("id", "")
                st.info(f"影片標題：**{video_title}**")
            except Exception as e:
                st.error(f"無法取得影片資訊：{e}"); st.stop()

        # Step 2: Thumbnail
        thumb_bytes: bytes | None = None
        with st.spinner("正在下載縮圖…"):
            thumb_bytes = get_thumbnail(video_id)
            if thumb_bytes: st.success("縮圖下載完成")

        # Step 3: Subtitles
        with st.spinner("正在下載韓文字幕…"):
            sub_path, sub_info = download_subtitles(url, tmpdir)
            if not sub_path:
                st.error(f"字幕下載失敗：{sub_info}"); st.stop()
            st.success(f"字幕下載成功（{sub_info}）")

        # Step 4: Parse
        with st.spinner("正在解析字幕…"):
            try:
                entries = parse_vtt(sub_path) if sub_path.endswith(".vtt") else parse_srt(sub_path)
                if not entries:
                    st.error("字幕解析後沒有找到文字。"); st.stop()
                st.success(f"共解析到 {len(entries)} 句歌詞")
            except Exception as e:
                st.error(f"字幕解析失敗：{e}"); st.stop()

        # Step 5: Download video (optional)
        video_path: str | None = None
        frames: dict[float, bytes] = {}
        if use_mv_frames:
            dl_status = st.empty()
            with st.spinner("正在下載 MV 影片（720p）…"):
                video_path = download_video(url, tmpdir, dl_status)
                dl_status.empty()
                if video_path: st.success("影片下載完成")
                else: st.warning("影片下載失敗，改用縮圖背景")

        # Step 6: Extract frames
        if video_path:
            frame_status = st.empty()
            with st.spinner("正在批次擷取 MV 畫面…"):
                timestamps = [ts for ts, _ in entries]
                frames = extract_all_frames(video_path, timestamps, status_ph=frame_status)
                frame_status.empty()
                st.success(f"擷取 {len(frames)}/{len(entries)} 張 MV 畫面")

        # Step 7: Translate
        trans_status = st.empty()
        with st.spinner("正在翻譯成繁體中文…"):
            try:
                chinese_lines = translate_lines(entries, trans_status)
                trans_status.empty()
                failed = sum(1 for z in chinese_lines if not z)
                if failed: st.warning(f"翻譯完成（{failed} 句保留韓文原文）")
                else: st.success("翻譯完成！")
            except Exception as e:
                st.error(f"翻譯失敗：{e}"); st.stop()

        # Step 8: TTS
        tts_map: dict[int, bytes] = {}
        if use_tts:
            tts_status = st.empty()
            with st.spinner("正在生成韓文語音…"):
                total = len(entries)
                for i, (_, ko) in enumerate(entries):
                    tts_status.text(f"生成語音... ({i+1}/{total})")
                    audio = generate_tts(ko, voice="ko-KR-SunHiNeural")
                    if audio: tts_map[i] = audio
                tts_status.empty()
                st.success(f"語音生成完成（{len(tts_map)} 句）")

        # Step 9: Build HTML
        html_status = st.empty()
        with st.spinner("正在生成 HTML…"):
            try:
                html_content = build_html(
                    video_title, entries, chinese_lines,
                    thumb_bytes, frames, tts_map, html_status
                )
                html_bytes = html_content.encode("utf-8")
                html_status.empty()
                st.success(f"HTML 生成完成！共 {len(entries)+1} 頁（含封面）")
            except Exception as e:
                st.error(f"HTML 生成失敗：{e}"); st.stop()

    # Step 10: Download
    if html_bytes:
        safe = re.sub(r'[\\/*?:"<>|]', "_", video_title)[:50]
        st.download_button(
            label="⬇️ 下載投影片 (.html)",
            data=html_bytes,
            file_name=f"{safe}.html",
            mime="text/html",
        )

st.markdown("---")
st.markdown(
    "<p style='text-align:center;color:#888;font-size:0.8rem;'>"
    "僅供個人韓語學習用途 · For personal language learning only</p>",
    unsafe_allow_html=True,
)
