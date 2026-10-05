# 修改与维护

## 网页演示

首页为根目录 `index.html`，公共样式在 `assets/site.css`。两个原始互动页面是独立 HTML，样式与逻辑内嵌于各自文件。音频页面位于 `demos/audio-sampling-demo.html`，使用本地 WAV 和浏览器播放器。

所有资源路径均按当前文件的位置解析，所以网页既能通过 `file://` 打开，也能部署到 GitHub Pages 的项目子路径。

修改后运行：

```console
python tools/verify_project.py
```

然后通过浏览器核对桌面和手机布局，并手动测试滑块、波形切换、采样状态和音频播放。

## 音频素材重生成

正常播放无需依赖。重生成音频素材时需要 Python 3.10 或更新版本及以下包：

```console
python -m pip install -r tools/audio/requirements.txt
python tools/audio/build_audio_assets.py
```

默认会重新生成内置合成测试乐句。若需要换成自己的 WAV 片段：

```console
python tools/audio/build_audio_assets.py --source path/to/your-clip.wav --out-dir build/audio
```

输入支持标准 PCM WAV；输出仍是 44.1 kHz、单声道试听文件。先在 `build/audio` 中检查，再替换正式素材。音频图表优先使用 Windows 系统中文字体；其他系统如无相应字体，图中文字可能缺字，需要配置中文字体后再用于课堂展示。

音频处理采用 FFT 低通、线性重采样、归一化和边缘淡入淡出。原材料的数值结果不是理想无限长度 Sinc 重建。

## 音频课件重生成

需要 Windows 桌面 PowerPoint、PowerShell，以及安装了上述包的 Python。先关闭需要保留的未保存 PowerPoint 工作，再运行：

```powershell
powershell -File tools\audio\build_audio_sampling_demo.ps1
```

或指定自己的源文件与 Python：

```powershell
powershell -File tools\audio\build_audio_sampling_demo.ps1 -SourceWav .\your-clip.wav -PythonExe python
```

默认生成 `materials/slides/采样定理_音频采样率播放Demo_重生成.pptx`，原始课件保留。脚本通过 PowerPoint 自动化生成，并输出预览图到 `build/audio-preview`。这部分依赖具体 Office 环境；发布项目时未重新生成原始 PPTX。

## PowerPoint 互动源码

`tools/powerpoint/SamplingTheoremInteractive.bas` 是已有 PPTM 内互动演示模块的可检查源码，来自原始生成脚本。

现有 PPTM 可直接在支持 VBA 的 PowerPoint 中使用。模块供修改互动逻辑时参考；保留的源码不自动改动本机宏安全设置。原始生成器使用的机器配置操作没有作为项目启动步骤。

## GitHub Pages

`pages.yml` 工作流按 [GitHub 官方 Pages 工作流说明](https://docs.github.com/en/pages/getting-started-with-github-pages/using-custom-workflows-with-github-pages) 配置。向 `main` 分支提交后，先运行标准库检查，再导出网站并部署。

网站导出只包含首页、网页演示、素材和课堂文件。构建脚本、Git 配置和私有整理记录不会放入网站导出目录。

```console
python tools/export_site.py
```

输出位于 `build/site`。

## 有意替换原材料时

原材料哈希记录旨在防止遗漏和意外改变。修改文件后，应根据新的正式版本更新 `materials/manifest.json` 中相应的大小和 SHA-256，并在提交说明中注明。不要在链接或完整性检查报错时简单跳过检查。
