# HLM证书批量生成器

基于 Tauri + Vue 3 的桌面端证书批量生成工具，轻量、跨平台。

## 功能

- 导入证书模板图片（PNG/JPG/BMP）
- 导入 Excel 数据（.xlsx/.xls/.csv）
- 拖拽式字段位置编辑，方向键微调，Alt 吸附对齐
- 右侧浮动字段配置面板（字体、字号、颜色、加粗）
- 批量导出 PNG/JPEG/PDF，支持进度条和取消
- 文件名规则可视化编辑
- 预览缩放、全屏

## 开始使用

```bash
npm install
npm run tauri dev
```

## 构建

```bash
npm run tauri build
```

## 技术栈

- Tauri 2
- Vue 3
- Vite 6
- Rust
- pdf-lib、xlsx

## 对比 Electron 版

| | Electron | Tauri |
|--|---------|-------|
| 安装包体积 | ~150MB | ~5-10MB |
| 内存占用 | ~100-200MB | ~20-50MB |
| 后端语言 | Node.js | Rust |

Electron 版见 [hlm2](https://github.com/lianqi00/hlm2)。
