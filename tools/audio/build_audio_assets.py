from __future__ import annotations

import argparse
import json
import math
import wave
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFont


BASE_SR = 44100
MAX_SECONDS = 12.0


def _font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = [
        r"C:\Windows\Fonts\msyhbd.ttc" if bold else r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\simhei.ttf",
        r"C:\Windows\Fonts\arial.ttf",
    ]
    for path in candidates:
        if Path(path).exists():
            return ImageFont.truetype(path, size=size)
    return ImageFont.load_default()


def read_wav(path: Path) -> tuple[np.ndarray, int]:
    with wave.open(str(path), "rb") as wf:
        channels = wf.getnchannels()
        sr = wf.getframerate()
        width = wf.getsampwidth()
        raw = wf.readframes(wf.getnframes())

    if width == 1:
        data = np.frombuffer(raw, dtype=np.uint8).astype(np.float32)
        data = (data - 128.0) / 128.0
    elif width == 2:
        data = np.frombuffer(raw, dtype="<i2").astype(np.float32) / 32768.0
    elif width == 3:
        b = np.frombuffer(raw, dtype=np.uint8).reshape(-1, 3)
        signed = (
            b[:, 0].astype(np.int32)
            | (b[:, 1].astype(np.int32) << 8)
            | (b[:, 2].astype(np.int32) << 16)
        )
        signed = np.where(signed & 0x800000, signed - 0x1000000, signed)
        data = signed.astype(np.float32) / 8388608.0
    elif width == 4:
        data = np.frombuffer(raw, dtype="<i4").astype(np.float32) / 2147483648.0
    else:
        raise ValueError(f"Unsupported WAV sample width: {width}")

    if channels > 1:
        data = data.reshape(-1, channels).mean(axis=1)
    return data.astype(np.float32), sr


def write_wav(path: Path, samples: np.ndarray, sr: int = BASE_SR) -> None:
    x = np.asarray(samples, dtype=np.float32)
    peak = float(np.max(np.abs(x))) if x.size else 1.0
    if peak > 0.98:
        x = x / peak * 0.98
    pcm = np.clip(x, -1.0, 1.0)
    pcm = (pcm * 32767.0).astype("<i2")
    path.parent.mkdir(parents=True, exist_ok=True)
    with wave.open(str(path), "wb") as wf:
        wf.setnchannels(1)
        wf.setsampwidth(2)
        wf.setframerate(sr)
        wf.writeframes(pcm.tobytes())


def resample_linear(x: np.ndarray, src_sr: int, dst_sr: int) -> np.ndarray:
    if src_sr == dst_sr:
        return x.astype(np.float32, copy=True)
    if x.size == 0:
        return x.astype(np.float32)
    dst_len = max(1, int(round(x.size * dst_sr / src_sr)))
    src_pos = np.arange(dst_len, dtype=np.float64) * (src_sr / dst_sr)
    return np.interp(src_pos, np.arange(x.size, dtype=np.float64), x).astype(np.float32)


def lowpass_fft(x: np.ndarray, sr: int, cutoff_hz: float) -> np.ndarray:
    cutoff_hz = min(cutoff_hz, sr * 0.49)
    if cutoff_hz <= 0:
        return np.zeros_like(x)
    spectrum = np.fft.rfft(x)
    freqs = np.fft.rfftfreq(x.size, 1.0 / sr)
    transition = max(120.0, cutoff_hz * 0.08)
    mask = np.ones_like(freqs)
    roll = (freqs - cutoff_hz) / transition
    mask = 1.0 / (1.0 + np.exp(np.clip(roll * 10.0, -60.0, 60.0)))
    filtered = np.fft.irfft(spectrum * mask, n=x.size)
    return filtered.astype(np.float32)


def fade_edges(x: np.ndarray, sr: int, fade_s: float = 0.05) -> np.ndarray:
    y = x.astype(np.float32, copy=True)
    n = min(int(sr * fade_s), y.size // 2)
    if n > 1:
        ramp = np.linspace(0.0, 1.0, n, dtype=np.float32)
        y[:n] *= ramp
        y[-n:] *= ramp[::-1]
    return y


def normalize(x: np.ndarray, target: float = 0.88) -> np.ndarray:
    peak = float(np.max(np.abs(x))) if x.size else 0.0
    if peak < 1e-9:
        return x.astype(np.float32)
    return (x / peak * target).astype(np.float32)


def synthetic_music(sr: int = BASE_SR, seconds: float = 10.0) -> np.ndarray:
    n = int(sr * seconds)
    x = np.zeros(n, dtype=np.float32)
    bpm = 92.0
    beat = 60.0 / bpm
    notes = [
        261.63,
        329.63,
        392.00,
        493.88,
        440.00,
        392.00,
        329.63,
        293.66,
        349.23,
        440.00,
        523.25,
        659.25,
        587.33,
        523.25,
        440.00,
        392.00,
    ]
    for i, freq in enumerate(notes * 2):
        start = int(i * beat * 0.75 * sr)
        dur = int(beat * 0.68 * sr)
        if start >= n:
            break
        end = min(n, start + dur)
        t = np.arange(end - start, dtype=np.float32) / sr
        env = np.ones_like(t)
        attack = max(1, int(0.025 * sr))
        release = max(1, int(0.14 * sr))
        env[: min(attack, env.size)] *= np.linspace(0, 1, min(attack, env.size))
        env[-min(release, env.size) :] *= np.linspace(1, 0, min(release, env.size))
        wave_note = (
            np.sin(2 * np.pi * freq * t)
            + 0.34 * np.sin(2 * np.pi * 2 * freq * t)
            + 0.20 * np.sin(2 * np.pi * 3 * freq * t)
            + 0.10 * np.sin(2 * np.pi * 5 * freq * t)
        )
        x[start:end] += (0.16 * env * wave_note).astype(np.float32)

    t_all = np.arange(n, dtype=np.float32) / sr
    bass = 0.13 * np.sin(2 * np.pi * 65.41 * t_all) * (
        0.55 + 0.45 * np.sin(2 * np.pi * (1.0 / (beat * 2.0)) * t_all)
    )
    shimmer = 0.035 * np.sin(2 * np.pi * 11800 * t_all) * (
        0.5 + 0.5 * np.sin(2 * np.pi * 5.1 * t_all)
    )
    x += bass.astype(np.float32)
    x += shimmer.astype(np.float32)

    rng = np.random.default_rng(20260531)
    for i in range(int(seconds / (beat / 2))):
        start = int(i * beat * 0.5 * sr)
        dur = int(0.055 * sr)
        end = min(n, start + dur)
        if end <= start:
            continue
        noise = rng.normal(0, 1, end - start).astype(np.float32)
        decay = np.exp(-np.linspace(0, 6.0, end - start)).astype(np.float32)
        carrier = np.sin(2 * np.pi * 9300 * np.arange(end - start) / sr).astype(np.float32)
        x[start:end] += 0.055 * noise * carrier * decay

    return fade_edges(normalize(x), sr, 0.08)


def prepare_source(source: Path | None, out_dir: Path) -> tuple[np.ndarray, str]:
    if source:
        x, sr = read_wav(source)
        x = x[: int(sr * MAX_SECONDS)]
        x = resample_linear(x, sr, BASE_SR)
        label = f"授权音频：{source.name}"
    else:
        x = synthetic_music(BASE_SR, 10.5)
        label = "内置合成测试乐句"
    x = fade_edges(normalize(x), BASE_SR, 0.08)
    write_wav(out_dir / "source_reference_44100.wav", x, BASE_SR)
    return x, label


def make_clip(x: np.ndarray, target_sr: int, anti_alias: bool = True) -> np.ndarray:
    if target_sr == BASE_SR:
        return x.copy()
    y = x
    if anti_alias:
        cutoff = min(target_sr * 0.46, BASE_SR * 0.48)
        y = lowpass_fft(y, BASE_SR, cutoff)
    captured = resample_linear(y, BASE_SR, target_sr)
    reconstructed = resample_linear(captured, target_sr, BASE_SR)
    if reconstructed.size > x.size:
        reconstructed = reconstructed[: x.size]
    elif reconstructed.size < x.size:
        reconstructed = np.pad(reconstructed, (0, x.size - reconstructed.size))
    return fade_edges(normalize(reconstructed, 0.88), BASE_SR, 0.03)


def spectrum_curve(x: np.ndarray, sr: int = BASE_SR, max_hz: int = 20000) -> tuple[np.ndarray, np.ndarray]:
    seg = x[: min(x.size, sr * 6)]
    window = np.hanning(seg.size)
    mag = np.abs(np.fft.rfft(seg * window))
    freqs = np.fft.rfftfreq(seg.size, 1.0 / sr)
    keep = freqs <= max_hz
    freqs = freqs[keep]
    mag = mag[keep]
    bins = np.linspace(0, max_hz, 360)
    values = np.zeros(len(bins) - 1)
    for i in range(len(values)):
        mask = (freqs >= bins[i]) & (freqs < bins[i + 1])
        values[i] = np.mean(mag[mask]) if np.any(mask) else 0
    values = np.log10(values + 1e-7)
    values -= values.min()
    if values.max() > 0:
        values /= values.max()
    centers = (bins[:-1] + bins[1:]) / 2
    return centers, values


def draw_spectra(clips: list[dict], out_path: Path) -> None:
    width, height = 1920, 1080
    img = Image.new("RGB", (width, height), "#F7F8F4")
    draw = ImageDraw.Draw(img)
    title_font = _font(54, True)
    body_font = _font(27)
    small_font = _font(22)
    draw.text((88, 58), "采样率降低时，高频细节先被切掉", font=title_font, fill="#20242A")
    draw.text(
        (90, 126),
        "纵向越高表示该频段能量越强；低采样率为了避免混叠，必须在 Nyquist 频率前留出滤波过渡带。",
        font=body_font,
        fill="#59616C",
    )
    left, right = 260, 1800
    top, row_h = 210, 128
    colors = ["#0B6E69", "#BB6B00", "#7B4FB4", "#D14256", "#6A737D"]
    for idx, clip in enumerate(clips):
        y0 = top + idx * row_h
        draw.text((88, y0 + 14), clip["short_label"], font=body_font, fill="#20242A")
        draw.text((88, y0 + 52), f"Nyquist {clip['nyquist_khz']} kHz", font=small_font, fill="#69717B")
        draw.line((left, y0 + 92, right, y0 + 92), fill="#D5D9D4", width=2)
        for hz in [4000, 8000, 12000, 16000, 20000]:
            x_axis = left + int((hz / 20000) * (right - left))
            draw.line((x_axis, y0 + 14, x_axis, y0 + 96), fill="#E7E9E4", width=1)
            if idx == len(clips) - 1:
                draw.text((x_axis - 28, y0 + 104), f"{hz//1000}k", font=small_font, fill="#828891")
        freqs = np.array(clip["spectrum_freqs"], dtype=np.float32)
        values = np.array(clip["spectrum_values"], dtype=np.float32)
        pts = []
        for f, v in zip(freqs, values):
            px = left + int((float(f) / 20000.0) * (right - left))
            py = y0 + 92 - int(float(v) * 76)
            pts.append((px, py))
        if len(pts) > 1:
            draw.line(pts, fill=colors[idx % len(colors)], width=4, joint="curve")
        cutoff_x = left + int((clip["effective_cutoff_hz"] / 20000) * (right - left))
        draw.line((cutoff_x, y0 + 10, cutoff_x, y0 + 96), fill="#111111", width=2)
    img.save(out_path)


def draw_waveform(clips: list[dict], out_path: Path) -> None:
    width, height = 1920, 1080
    img = Image.new("RGB", (width, height), "#FFFFFF")
    draw = ImageDraw.Draw(img)
    title_font = _font(52, True)
    body_font = _font(27)
    small_font = _font(21)
    draw.text((88, 58), "同一声音，用更稀疏的点记录", font=title_font, fill="#20242A")
    draw.text(
        (90, 125),
        "图中放大了 20 ms 的波形。采样点变少后，重建曲线只能保留较慢变化，细节逐渐消失。",
        font=body_font,
        fill="#59616C",
    )
    panel_w, panel_h = 540, 260
    positions = [(90, 220), (690, 220), (1290, 220), (390, 560), (990, 560)]
    colors = ["#0B6E69", "#BB6B00", "#7B4FB4", "#D14256", "#6A737D"]
    for i, clip in enumerate(clips):
        x0, y0 = positions[i]
        draw.rounded_rectangle((x0, y0, x0 + panel_w, y0 + panel_h), radius=18, fill="#F5F6F2", outline="#DADDD5", width=2)
        draw.text((x0 + 26, y0 + 18), clip["short_label"], font=body_font, fill="#20242A")
        draw.text((x0 + 26, y0 + 55), f"每秒 {clip['target_rate']:,} 个采样点", font=small_font, fill="#69717B")
        line_y = y0 + 158
        draw.line((x0 + 28, line_y, x0 + panel_w - 28, line_y), fill="#D5D9D4", width=2)
        data = np.array(clip["waveform"], dtype=np.float32)
        if data.size:
            pts = []
            for j, val in enumerate(data):
                px = x0 + 32 + int(j / (data.size - 1) * (panel_w - 64))
                py = line_y - int(float(val) * 72)
                pts.append((px, py))
            draw.line(pts, fill=colors[i % len(colors)], width=4)
        dots = np.array(clip["sample_dots"], dtype=np.float32)
        for sx, sy in dots:
            px = x0 + 32 + int(float(sx) * (panel_w - 64))
            py = line_y - int(float(sy) * 72)
            draw.ellipse((px - 5, py - 5, px + 5, py + 5), fill="#20242A")
    img.save(out_path)


def build_assets(source: Path | None, out_dir: Path) -> dict:
    out_dir.mkdir(parents=True, exist_ok=True)
    x, source_label = prepare_source(source, out_dir)

    clip_specs = [
        ("clip_44100_reference.wav", "44.1 kHz", "CD/常规音乐", 44100, True),
        ("clip_22050.wav", "22.05 kHz", "高频明显变少", 22050, True),
        ("clip_11025.wav", "11.025 kHz", "质感变薄", 11025, True),
        ("clip_8000.wav", "8 kHz", "电话感", 8000, True),
        ("clip_4000.wav", "4 kHz", "严重失真", 4000, True),
        ("clip_8000_alias.wav", "8 kHz", "无抗混叠示例", 8000, False),
    ]
    clips: list[dict] = []
    for filename, label, note, target_sr, anti_alias in clip_specs:
        y = make_clip(x, target_sr, anti_alias)
        write_wav(out_dir / filename, y, BASE_SR)
        freqs, values = spectrum_curve(y)
        start = min(int(BASE_SR * 2.0), max(0, y.size - int(BASE_SR * 0.03)))
        zoom = y[start : start + int(BASE_SR * 0.020)]
        zoom_ds = resample_linear(zoom, BASE_SR, 220)
        dot_count = max(5, int(round(target_sr / BASE_SR * 56)))
        dot_x = np.linspace(0, 1, dot_count)
        dot_y = np.interp(dot_x, np.linspace(0, 1, zoom_ds.size), zoom_ds) if zoom_ds.size else np.zeros(dot_count)
        cutoff = BASE_SR / 2 if target_sr == BASE_SR else (target_sr * 0.46 if anti_alias else target_sr / 2)
        clips.append(
            {
                "file": filename,
                "filename": filename,
                "short_label": label,
                "note": note,
                "target_rate": target_sr,
                "nyquist_khz": round(target_sr / 2000.0, 2),
                "anti_alias": anti_alias,
                "effective_cutoff_hz": float(min(cutoff, 20000)),
                "spectrum_freqs": freqs.astype(float).round(1).tolist(),
                "spectrum_values": values.astype(float).round(4).tolist(),
                "waveform": zoom_ds.astype(float).round(4).tolist(),
                "sample_dots": np.column_stack([dot_x, dot_y]).astype(float).round(4).tolist(),
            }
        )

    main_clips = clips[:5]
    spectra_path = out_dir / "spectra_comparison.png"
    waveform_path = out_dir / "waveform_zoom.png"
    draw_spectra(main_clips, spectra_path)
    draw_waveform(main_clips, waveform_path)

    manifest = {
        "base_sample_rate": BASE_SR,
        "source_label": source_label,
        "source_file": source.name if source else "",
        "clips": clips,
        "main_clips": main_clips,
        "spectra_image": spectra_path.name,
        "waveform_image": waveform_path.name,
    }
    with (out_dir / "manifest.json").open("w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    return manifest


def main() -> None:
    parser = argparse.ArgumentParser(description="Build audio assets for a sampling-rate PowerPoint demo.")
    parser.add_argument("--source", type=Path, default=None, help="Optional licensed mono/stereo WAV source.")
    parser.add_argument("--out-dir", type=Path, default=Path(__file__).resolve().parents[2] / "assets" / "audio")
    args = parser.parse_args()
    if args.source and not args.source.exists():
        raise FileNotFoundError(args.source)
    manifest = build_assets(args.source, args.out_dir)
    print(json.dumps({"manifest": str((args.out_dir / "manifest.json").resolve()), "clips": len(manifest["clips"])}, ensure_ascii=False))


if __name__ == "__main__":
    main()
