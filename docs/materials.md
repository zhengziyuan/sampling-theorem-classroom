# 材料目录

## 原始课堂文件

项目收录 5 份课件和 2 种讲义格式。文件保持原始内容，原文件名便于与已有授课记录对应。

| 文件 | 内容 |
| --- | --- |
| [15 分钟课堂 PPTX](../materials/slides/采样定理_15分钟课堂.pptx) | 11 页，采样、频谱复制、混叠、重建、工程例题和小测 |
| [带互动演示的 PPTM](../materials/slides/采样定理_15分钟课堂_含互动演示器.pptm) | 12 页，最后一页包含 VBA 互动按钮 |
| [HTML 演示器插入版](../materials/slides/采样定理_HTML演示器_PPT插入版.pptx) | 4 页，正常采样、混叠折回和重建公式画面 |
| [图片采样扩展版](../materials/slides/采样定理_图片采样Demo_扩展版.pptx) | 9 页，前述画面加二维图片采样实验 |
| [音频采样率播放 Demo](../materials/slides/采样定理_音频采样率播放Demo.pptx) | 7 页，内嵌合成测试音频 |
| [PDF 讲义](../materials/handouts/采样定理_15分钟讲义.pdf) | 6 页，适合阅读与打印 |
| [Word 讲义](../materials/handouts/采样定理_15分钟讲义.docx) | 可编辑讲义 |

## 网页与素材

两个原始网页演示位于 [demos](../demos)，新增音频试听页使用已有 WAV 素材。

- [PPT 插入图](../assets/ppt)：3 张 1920×1080 图像，文件名沿用原始编号。
- [讲义图示](../assets/figures)：采样、频谱复制、混叠波形和工程链路。
- [音频素材](../assets/audio)：参考音频、6 个试听文件、频谱与波形图、相对路径元数据。
- [截图预览](../assets/screenshots)：原有演示截图与课件预览，供项目首页和 README 展示。

预览截图来自现有教学材料的保存输出。它们是静态展示图，交互状态以现场网页为准。

## 完整性记录

[materials/manifest.json](../materials/manifest.json) 记录原材料的大小、SHA-256、课件页数及是否含 VBA。

`python tools/verify_project.py` 会核对原文件、项目内部链接、Office 压缩包和 WAV 参数。整理后的音频 JSON 去掉了原始个人电脑路径。原始音频课件印有旧版重生成命令，项目保留了兼容入口 `outputs/audio-sampling-demo/build_audio_sampling_demo.ps1`。
